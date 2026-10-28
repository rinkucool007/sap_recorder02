"""
SAP GUI Recorder - desktop app to record SAP ECC (SAP GUI for Windows) steps and replay them.

    pip install pywin32
    python sap_recorder_ui.py

1. Log on to SAP, pick the session in the dropdown (Refresh if needed).
2. Click  Record  and work in SAP GUI. Every click / typed value appears as a step.
3. Click  Stop.  The Python replay code is generated in the "Python code" tab.
4. Click  Play  - the generated code is saved and executed; output shows in the Log tab.
5. Export .py to keep the script, Save to keep the steps (JSON) for later editing.

Everything uses the SAP GUI Scripting API over COM. Calls go through raw IDispatch.Invoke,
so they work whether or not makepy / early binding has been generated (name case does not matter).
"""
from __future__ import annotations

import datetime as _dt
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_TITLE = "SAP GUI Recorder"
WORK_DIR = Path.home() / "SAPRecorder"
ID_PREFIX_RE = re.compile(r"^/app/con\[\d+\]/ses\[\d+\]/")
SES_ID_RE = re.compile(r"/app/con\[(\d+)\]/ses\[(\d+)\]")
PARAM_RE = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")
SKIP_PROPS = {"caretposition"}
SKIP_METHODS = {"resizeworkingpane"}
ROUND_TRIP = {"sendvkey", "press", "presstoolbarbutton", "pressbutton", "doubleclickcurrentcell",
              "doubleclicknode", "doubleclickitem", "clickcurrentcell", "selectcontextmenuitem",
              "presstoolbarcontextbutton", "selectitem", "select"}


# ======================================================================================
# Pure logic: turning SAP Change events into steps, generating Python code
# ======================================================================================
def short_id(full_id: str) -> str:
    return ID_PREFIX_RE.sub("", full_id or "")


def to_py(v: Any) -> Any:
    if isinstance(v, (list, tuple)):
        return [to_py(x) for x in v]
    if isinstance(v, (str, int, float, bool)) or v is None:
        return v
    return str(v)


def normalise_command_array(arr: Any) -> list[list[Any]]:
    arr = to_py(arr)
    if not isinstance(arr, list) or not arr:
        return []
    if isinstance(arr[0], list):
        return [c for c in arr if isinstance(c, list) and c]
    return [arr]


def command_to_step(component_id: str, cmd: list[Any]) -> dict | None:
    kind = str(cmd[0]).upper()
    sid = short_id(component_id)
    if kind == "SP" and len(cmd) >= 3:
        return {"id": sid, "type": "SP", "name": str(cmd[1]), "value": cmd[2]}
    if kind == "M" and len(cmd) >= 2:
        return {"id": sid, "type": "M", "name": str(cmd[1]), "args": list(cmd[2:])}
    return None


class StepCollector:
    def __init__(self, steps: list[dict] | None = None):
        self.steps: list[dict] = steps if steps is not None else []

    def add_event(self, component_id: str, command_array: Any, tcode: str = "") -> int:
        """Returns how many visible changes were made (new or merged steps)."""
        changed = 0
        for cmd in normalise_command_array(command_array):
            step = command_to_step(component_id, cmd)
            if not step:
                continue
            nm = step["name"].lower()
            if step["type"] == "SP" and nm in SKIP_PROPS:
                continue
            if step["type"] == "M" and nm in SKIP_METHODS:
                continue
            if tcode:
                step["tcode"] = tcode
            last = self.steps[-1] if self.steps else None
            if (last and not last.get("disabled") and step["type"] == "SP" and last["type"] == "SP"
                    and last["id"] == step["id"] and last["name"].lower() == nm):
                last["value"] = step["value"]          # typing in same field -> keep final value
            else:
                self.steps.append(step)
            changed += 1
        return changed


def describe(step: dict) -> tuple[str, str]:
    if step["type"] == "SP":
        return "Set", repr(step["value"])
    return "Call", ", ".join(repr(a) for a in step.get("args", []))


def _py_value(v: Any) -> str:
    if isinstance(v, str) and PARAM_RE.search(v):
        full = PARAM_RE.fullmatch(v)
        if full:
            return f"DATA[{full.group(1)!r}]"
        out = []
        for i, part in enumerate(PARAM_RE.split(v)):
            if i % 2 == 0:
                if part:
                    out.append(repr(part))
            else:
                out.append(f"str(DATA[{part!r}])")
        return " + ".join(out)
    return repr(v)


def find_params(steps: list[dict]) -> list[str]:
    found: list[str] = []
    for s in steps:
        vals = [s.get("value")] + list(s.get("args", []))
        for v in vals:
            if isinstance(v, str):
                for m in PARAM_RE.findall(v):
                    if m not in found:
                        found.append(m)
    return found


REQUIREMENTS_TXT = """\
# SAP GUI Recorder / generated replay scripts
# Windows only - SAP GUI Scripting is a Windows COM API.
# tkinter (the UI) ships with the standard Python installer from python.org.
pywin32>=306; sys_platform == "win32"
"""


def write_requirements(folder: Path) -> Path:
    req = Path(folder) / "requirements.txt"
    if not req.exists():
        req.write_text(REQUIREMENTS_TXT, encoding="utf-8")
    return req


RUNTIME = r'''
import argparse
import os
import sys
import time

try:
    import pythoncom
    import win32com.client
except ImportError:
    sys.exit("pywin32 is required:  pip install pywin32")

_GET = pythoncom.DISPATCH_PROPERTYGET | pythoncom.DISPATCH_METHOD
_CALL = pythoncom.DISPATCH_METHOD
_PUT = pythoncom.DISPATCH_PROPERTYPUT
ROUND_TRIP = __ROUND_TRIP__
OPTIONS = {"delay": 0.0, "stop_on_error": True, "screenshots": "replay_screenshots"}
STATE = {"step": 0}


class ReplayError(Exception):
    pass


def log(msg):
    print(msg, flush=True)


# --- late-bound COM calls: name case does not matter, works with or without makepy ---
def _ole(obj):
    return getattr(obj, "_oleobj_", obj)


def com_get(obj, name, *args):
    o = _ole(obj)
    return o.Invoke(o.GetIDsOfNames(name), 0, _GET, True, *args)


def com_call(obj, name, *args):
    o = _ole(obj)
    return o.Invoke(o.GetIDsOfNames(name), 0, _CALL, True, *args)


def com_put(obj, name, value):
    o = _ole(obj)
    o.Invoke(o.GetIDsOfNames(name), 0, _PUT, False, value)


def com_error_text(exc):
    try:
        info = exc.excepinfo
        if info and info[2]:
            return str(info[2]).strip()
    except Exception:
        pass
    return str(exc)


# --- SAP helpers ---
def connect(conn=0, ses=0):
    try:
        sapgui = win32com.client.GetObject("SAPGUI")
    except Exception as exc:
        raise ReplayError("SAP Logon is not running (GetObject('SAPGUI') failed): %s" % exc)
    app = com_get(sapgui, "GetScriptingEngine")
    if app is None:
        raise ReplayError("SAP GUI scripting engine not available - enable scripting in SAP GUI options")
    session = com_call(app, "findById", "/app/con[%d]/ses[%d]" % (conn, ses), False)
    if session is None:
        raise ReplayError("SAP session con[%d]/ses[%d] not found - log on first" % (conn, ses))
    return session


def wait_idle(session, timeout=60):
    end = time.time() + timeout
    while com_get(session, "Busy"):
        if time.time() > end:
            raise ReplayError("SAP still busy after %ss" % timeout)
        time.sleep(0.1)


def find(session, cid, timeout=10):
    end = time.time() + timeout
    while True:
        try:
            obj = com_call(session, "findById", cid, False)
        except pythoncom.com_error:
            obj = None
        if obj is not None:
            return obj
        if time.time() > end:
            raise ReplayError("Control not found on screen: " + cid)
        time.sleep(0.25)


def _target(obj, name):
    parts = name.split(".")              # e.g. verticalScrollbar.position
    for p in parts[:-1]:
        obj = com_get(obj, p)
    return obj, parts[-1]


def status_bar(session):
    try:
        sbar = com_call(session, "findById", "wnd[0]/sbar", False)
        if sbar is None:
            return "", ""
        return (com_get(sbar, "MessageType") or ""), (com_get(sbar, "Text") or "")
    except pythoncom.com_error:
        return "", ""


def screenshot(session, n):
    try:
        os.makedirs(OPTIONS["screenshots"], exist_ok=True)
        path = os.path.abspath(os.path.join(OPTIONS["screenshots"], "step%03d_%d.png" % (n, int(time.time()))))
        com_call(com_call(session, "findById", "wnd[0]"), "HardCopy", path, 2)
        log("     screenshot: " + path)
    except Exception as exc:
        log("     (screenshot failed: %s)" % exc)


def _pause():
    if OPTIONS["delay"]:
        time.sleep(OPTIONS["delay"])


def set_value(session, n, cid, name, value):
    STATE["step"] = n
    log("%3d  SET   %s.%s = %r" % (n, cid, name, value))
    wait_idle(session)
    obj, prop = _target(find(session, cid), name)
    com_put(obj, prop, value)
    _pause()


def call_method(session, n, cid, name, *args):
    STATE["step"] = n
    log("%3d  CALL  %s.%s(%s)" % (n, cid, name, ", ".join(repr(a) for a in args)))
    wait_idle(session)
    obj, meth = _target(find(session, cid), name)
    com_call(obj, meth, *args)
    if meth.lower() in ROUND_TRIP:
        wait_idle(session)
        mtype, text = status_bar(session)
        if text:
            log("       [%s] %s" % (mtype or "-", text))
        if mtype in ("E", "A") and OPTIONS["stop_on_error"]:
            raise ReplayError("SAP %s message: %s" % ("error" if mtype == "E" else "abort", text))
    _pause()
'''

MAIN = r'''

def main():
    ap = argparse.ArgumentParser(description="Replay recorded SAP GUI steps")
    ap.add_argument("--conn", type=int, default=0, help="connection index (con[N])")
    ap.add_argument("--ses", type=int, default=0, help="session index (ses[N])")
    ap.add_argument("--delay", type=float, default=0.0, help="seconds to wait between steps")
    ap.add_argument("--keep-going", action="store_true", help="do not stop on SAP error messages")
    ap.add_argument("--screenshots", default="replay_screenshots")
    a = ap.parse_args()
    OPTIONS.update(delay=a.delay, stop_on_error=not a.keep_going, screenshots=a.screenshots)

    pythoncom.CoInitialize()
    try:
        session = connect(a.conn, a.ses)
    except Exception as exc:
        log("CONNECT FAILED: %s" % (com_error_text(exc) if isinstance(exc, pythoncom.com_error) else exc))
        return 2
    try:
        info = com_get(session, "Info")
        log("Connected: system %s, client %s, user %s" % (
            com_get(info, "SystemName"), com_get(info, "Client"), com_get(info, "User")))
    except Exception:
        pass

    start = time.time()
    try:
        run(session)
    except Exception as exc:
        msg = com_error_text(exc) if isinstance(exc, pythoncom.com_error) else str(exc)
        log("FAILED at step %d: %s" % (STATE["step"], msg))
        screenshot(session, STATE["step"])
        log("RESULT: FAIL")
        return 1
    log("RESULT: PASS  (%d steps in %.1fs)" % (STATE["step"], time.time() - start))
    return 0


if __name__ == "__main__":
    sys.exit(main())
'''


def generate_python(steps: list[dict], meta: dict | None = None) -> str:
    meta = meta or {}
    active = [s for s in steps if not s.get("disabled")]
    params = find_params(active)
    head = [
        '"""',
        "SAP GUI replay script - generated by SAP GUI Recorder on "
        + _dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        f"Recorded on: system {meta.get('system', '?')}, client {meta.get('client', '?')}, "
        f"user {meta.get('user', '?')}",
        "",
        "Run:   python this_file.py [--conn 0 --ses 0] [--delay 0.5] [--keep-going]",
        "Needs: pip install -r requirements.txt (pywin32), SAP Logon open and logged on, scripting enabled.",
        '"""',
    ]
    code = "\n".join(head) + "\n" + RUNTIME.replace("__ROUND_TRIP__", repr(sorted(ROUND_TRIP)))
    if params:
        code += "\n\n# Values for ${...} placeholders - edit here\nDATA = {\n"
        code += "".join(f"    {p!r}: \"\",\n" for p in params) + "}\n"
    body = ["", "", "def run(s):", '    """Recorded steps."""']
    last_tcode = None
    for n, st in enumerate(active, 1):
        tc = st.get("tcode")
        if tc and tc != last_tcode:
            body.append(f"    # --- transaction {tc} ---")
            last_tcode = tc
        if st["type"] == "SP":
            body.append(f"    set_value(s, {n}, {st['id']!r}, {st['name']!r}, {_py_value(st['value'])})")
        else:
            args = "".join(", " + _py_value(a) for a in st.get("args", []))
            body.append(f"    call_method(s, {n}, {st['id']!r}, {st['name']!r}{args})")
    if not active:
        body.append("    pass  # nothing recorded yet")
    return code + "\n".join(body) + "\n" + MAIN


# ======================================================================================
# COM layer used by the UI (recording, session list, diagnostics)
# ======================================================================================
_rt: dict = {}


def runtime():
    """Executes the same helper functions the generated scripts use, so UI and scripts behave alike."""
    if not _rt:
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
        except ImportError as exc:
            raise RuntimeError("pywin32 is not installed (pip install pywin32), or this is not Windows.\n"
                               + str(exc))
        exec(RUNTIME.replace("__ROUND_TRIP__", repr(sorted(ROUND_TRIP))), _rt)
    return _rt


def _elements(coll):
    rt = runtime()
    n = rt["com_get"](coll, "Count")
    for i in range(n):
        try:
            yield rt["com_call"](coll, "ElementAt", i)
        except Exception:
            yield rt["com_call"](coll, "Item", i)


def list_sessions() -> list[dict]:
    rt = runtime()
    import win32com.client
    sapgui = win32com.client.GetObject("SAPGUI")
    app = rt["com_get"](sapgui, "GetScriptingEngine")
    out = []
    for conn in _elements(rt["com_get"](app, "Children")):
        disabled = False
        try:
            disabled = bool(rt["com_get"](conn, "DisabledByServer"))
        except Exception:
            pass
        for ses in _elements(rt["com_get"](conn, "Children")):
            sid = rt["com_get"](ses, "Id")
            m = SES_ID_RE.search(sid or "")
            if not m:
                continue
            info = rt["com_get"](ses, "Info")
            g = lambda k: rt["com_get"](info, k)  # noqa: E731
            label = f"con[{m.group(1)}]/ses[{m.group(2)}]  {g('SystemName')} · {g('Client')} · {g('User')} · {g('Transaction')}"
            if disabled:
                label += "  (SCRIPTING DISABLED BY SERVER)"
            out.append({"conn": int(m.group(1)), "ses": int(m.group(2)), "label": label, "disabled": disabled})
    return out


def attach_events(session_ole, handler_cls):
    """Subscribe handler_cls.OnChange to the session's events. Tries pywin32's WithEvents,
    then builds the event class from SAP's type library directly."""
    import pythoncom  # noqa: F401
    import win32com.client
    from win32com.client import dynamic, gencache
    errors = []
    try:
        return win32com.client.WithEvents(session_ole, handler_cls)
    except Exception as exc:
        errors.append(f"WithEvents: {exc}")
    try:
        ole = getattr(session_ole, "_oleobj_", session_ole)
        tlb, _ = ole.GetTypeInfo().GetContainingTypeLib()
        la = tlb.GetLibAttr()
        mod = gencache.EnsureModule(str(la[0]), la[1], la[3], la[4], bValidateFile=0)
        src = None
        for name in dir(mod):
            obj = getattr(mod, name)
            if isinstance(obj, type) and getattr(obj, "default_source", None) and "session" in name.lower():
                src = obj.default_source
                break
        if src is None:
            raise RuntimeError("no GuiSession event interface in type library")
        sink_cls = type("SapSessionSink", (src, handler_cls), {})
        return sink_cls(dynamic.Dispatch(ole))
    except Exception as exc:
        errors.append(f"type library: {exc}")
    raise RuntimeError("Could not subscribe to SAP events.\n" + "\n".join(errors)
                       + "\nTry once:  python -m win32com.client.makepy  -> 'SAP GUI Scripting API'")


class RecorderThread(threading.Thread):
    """Owns its own COM apartment: attaches, sets Record=True, pumps messages, forwards events."""

    def __init__(self, conn: int, ses: int, out: queue.Queue):
        super().__init__(daemon=True)
        self.conn, self.ses, self.out = conn, ses, out
        self.stop_evt = threading.Event()

    def stop(self):
        self.stop_evt.set()

    def run(self):  # pragma: no cover - needs SAP
        import pythoncom
        pythoncom.CoInitialize()
        session, sink = None, None
        try:
            rt = runtime()
            session = rt["connect"](self.conn, self.ses)
            info = rt["com_get"](session, "Info")
            meta = {k.lower(): rt["com_get"](info, k) for k in ("SystemName", "Client", "User", "Language")}
            meta["system"] = meta.pop("systemname")
            self.out.put(("meta", meta))
            outer = self

            class Handler:
                def OnChange(self, Session, Component, CommandArray):  # noqa: N802,N803
                    try:
                        cid = rt["com_get"](Component, "Id")
                        try:
                            tcode = rt["com_get"](rt["com_get"](session, "Info"), "Transaction")
                        except Exception:
                            tcode = ""
                        outer.out.put(("change", cid, to_py(CommandArray), tcode))
                    except Exception as exc:
                        outer.out.put(("log", f"event not understood: {exc}"))

                def OnDestroy(self, *args):  # noqa: N802
                    outer.out.put(("log", "SAP session was closed - recording stopped."))
                    outer.stop_evt.set()

            sink = attach_events(session, Handler)
            rt["com_put"](session, "Record", True)
            self.out.put(("recording", None))
            while not self.stop_evt.is_set():
                pythoncom.PumpWaitingMessages()
                time.sleep(0.02)
            pythoncom.PumpWaitingMessages()       # deliver the last events
        except Exception as exc:
            self.out.put(("error", f"{exc}\n\n{traceback.format_exc()}"))
        finally:
            try:
                if session is not None:
                    runtime()["com_put"](session, "Record", False)
            except Exception:
                pass
            try:
                if sink is not None and hasattr(sink, "close"):
                    sink.close()
            except Exception:
                pass
            sink = session = None
            pythoncom.CoUninitialize()
            self.out.put(("stopped", None))


def diagnose(conn: int | None, ses: int | None) -> str:
    import platform
    import struct
    lines = [f"Python {platform.python_version()} ({struct.calcsize('P') * 8}-bit) on {platform.platform()}"]
    try:
        import win32api  # noqa: F401
        import win32com
        lines.append(f"pywin32 OK ({os.path.dirname(win32com.__file__)})")
    except Exception as exc:
        lines.append(f"pywin32 MISSING: {exc}  ->  pip install pywin32")
        return "\n".join(lines)
    try:
        import winreg
        key = r"Software\SAP\SAPGUI Front\SAP Frontend Server\Security"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as k:
            for name in ("UserScripting", "WarnOnAttach", "WarnOnConnection"):
                try:
                    lines.append(f"Client option {name} = {winreg.QueryValueEx(k, name)[0]}")
                except OSError:
                    lines.append(f"Client option {name} = (not set)")
    except OSError:
        lines.append("Client scripting options not found in registry (defaults apply)")
    try:
        sess = list_sessions()
        lines.append(f"SAP Logon reachable, {len(sess)} session(s):")
        lines += ["   " + s["label"] for s in sess]
    except Exception as exc:
        lines.append(f"SAP Logon NOT reachable: {exc}")
        lines.append("  -> start SAP Logon, log on, and enable scripting (Options > Accessibility & Scripting)")
        return "\n".join(lines)
    if conn is not None:
        import pythoncom
        try:
            rt = runtime()
            session = rt["connect"](conn, ses)

            class _H:
                def OnChange(self, *a):  # noqa: N802
                    pass
            sink = attach_events(session, _H)
            if hasattr(sink, "close"):
                sink.close()
            lines.append("Event subscription OK - recording will work")
            mtype, text = rt["status_bar"](session)
            lines.append(f"Status bar readable: [{mtype}] {text}")
        except Exception as exc:
            lines.append(f"Event subscription FAILED: {exc}")
        pythoncom.PumpWaitingMessages()
    return "\n".join(lines)


# ======================================================================================
# UI
# ======================================================================================
class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.q: queue.Queue = queue.Queue()
        self.collector = StepCollector()
        self.meta: dict = {}
        self.sessions: list[dict] = []
        self.rec: RecorderThread | None = None
        self.proc: subprocess.Popen | None = None
        self.json_path: Path | None = None
        WORK_DIR.mkdir(parents=True, exist_ok=True)

        root.title(APP_TITLE)
        root.geometry("1180x760")
        root.minsize(900, 560)
        self._style()
        self._build()
        self._refresh_buttons()
        self.refresh_sessions()
        root.after(100, self._poll)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- layout ----------
    def _style(self):
        st = ttk.Style(self.root)
        if "vista" in st.theme_names():
            st.theme_use("vista")
        elif "clam" in st.theme_names():
            st.theme_use("clam")
        st.configure("Rec.TButton", foreground="#B00020", font=("Segoe UI", 10, "bold"))
        st.configure("Play.TButton", foreground="#0A6E31", font=("Segoe UI", 10, "bold"))
        st.configure("Stop.TButton", font=("Segoe UI", 10, "bold"))
        st.configure("Treeview", rowheight=22)

    def _build(self):
        r = self.root
        top = ttk.Frame(r, padding=(8, 8, 8, 4))
        top.pack(fill="x")
        ttk.Label(top, text="SAP session:").pack(side="left")
        self.sess_var = tk.StringVar()
        self.sess_cb = ttk.Combobox(top, textvariable=self.sess_var, state="readonly", width=70)
        self.sess_cb.pack(side="left", padx=6)
        ttk.Button(top, text="Refresh", command=self.refresh_sessions).pack(side="left")
        ttk.Button(top, text="Diagnose", command=self.run_diagnose).pack(side="left", padx=4)
        self.top_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Stay on top", variable=self.top_var,
                        command=lambda: r.attributes("-topmost", self.top_var.get())).pack(side="right")

        bar = ttk.Frame(r, padding=(8, 4))
        bar.pack(fill="x")
        self.btn_rec = ttk.Button(bar, text="●  Record", style="Rec.TButton", command=self.start_record)
        self.btn_stop = ttk.Button(bar, text="■  Stop", style="Stop.TButton", command=self.stop_all)
        self.btn_play = ttk.Button(bar, text="▶  Play", style="Play.TButton", command=self.play)
        for b in (self.btn_rec, self.btn_stop, self.btn_play):
            b.pack(side="left", padx=(0, 6), ipady=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="Open steps…", command=self.open_json).pack(side="left")
        ttk.Button(bar, text="Save steps…", command=self.save_json).pack(side="left", padx=4)
        ttk.Button(bar, text="Export .py…", command=self.export_py).pack(side="left")
        ttk.Button(bar, text="Clear", command=self.clear).pack(side="left", padx=4)

        opts = ttk.Frame(bar)
        opts.pack(side="right")
        self.stop_err = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Stop on SAP error", variable=self.stop_err).pack(side="left", padx=6)
        ttk.Label(opts, text="Delay (s):").pack(side="left")
        self.delay = tk.DoubleVar(value=0.0)
        ttk.Spinbox(opts, from_=0, to=5, increment=0.25, textvariable=self.delay, width=5).pack(side="left")

        paned = ttk.PanedWindow(r, orient="vertical")
        paned.pack(fill="both", expand=True, padx=8, pady=4)

        # steps table
        tf = ttk.Frame(paned)
        cols = ("n", "tcode", "action", "id", "name", "value")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="extended")
        for c, t, w, stretch in (("n", "#", 40, False), ("tcode", "Tcode", 125, False),
                                 ("action", "Action", 60, False), ("id", "Control ID", 430, True),
                                 ("name", "Property / Method", 150, False), ("value", "Value / Args", 260, True)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, stretch=stretch, anchor="w")
        self.tree.tag_configure("off", foreground="#9A9A9A")
        ysb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)
        sb = ttk.Frame(tf)
        sb.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        for text, cmd in (("Edit value", self.edit_step), ("Delete", self.delete_steps),
                          ("Enable / Disable", self.toggle_steps), ("Move up", lambda: self.move(-1)),
                          ("Move down", lambda: self.move(1))):
            ttk.Button(sb, text=text, command=cmd).pack(side="left", padx=(0, 4))
        ttk.Label(sb, text="Tip: start from SAP Easy Access and begin with /n<tcode>. "
                           "Edit a value to ${NAME} to make it a parameter.",
                  foreground="#555").pack(side="left", padx=10)
        self.tree.bind("<Double-1>", lambda e: self.edit_step())
        self.tree.bind("<Delete>", lambda e: self.delete_steps())
        paned.add(tf, weight=3)

        # code + log
        nb = ttk.Notebook(paned)
        self.code = self._text(nb, mono=True)
        self.log = self._text(nb, mono=True)
        nb.add(self.code.master, text="Python code (what Play runs)")
        nb.add(self.log.master, text="Log")
        self.nb = nb
        paned.add(nb, weight=2)

        self.status = tk.StringVar(value="Ready")
        ttk.Label(r, textvariable=self.status, anchor="w", padding=(10, 3), relief="sunken").pack(fill="x")
        self._update_code()

    def _text(self, parent, mono=False):
        f = ttk.Frame(parent)
        t = tk.Text(f, wrap="none", undo=True, font=("Consolas", 10) if mono else None)
        ys = ttk.Scrollbar(f, orient="vertical", command=t.yview)
        xs = ttk.Scrollbar(f, orient="horizontal", command=t.xview)
        t.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        t.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        f.rowconfigure(0, weight=1)
        f.columnconfigure(0, weight=1)
        return t

    # ---------- helpers ----------
    def _log(self, msg: str):
        self.log.insert("end", msg.rstrip("\n") + "\n")
        self.log.see("end")

    def _selected_session(self) -> dict | None:
        i = self.sess_cb.current()
        return self.sessions[i] if 0 <= i < len(self.sessions) else None

    def _refresh_buttons(self):
        recording = self.rec is not None
        playing = self.proc is not None
        self.btn_rec.state(["disabled"] if recording or playing else ["!disabled"])
        self.btn_play.state(["disabled"] if recording or playing else ["!disabled"])
        self.btn_stop.state(["!disabled"] if recording or playing else ["disabled"])

    def _refresh_tree(self, select_last=False):
        self.tree.delete(*self.tree.get_children())
        for i, s in enumerate(self.collector.steps):
            action, val = describe(s)
            self.tree.insert("", "end", iid=str(i), tags=("off",) if s.get("disabled") else (),
                             values=(i + 1, s.get("tcode", ""), action, s["id"], s["name"], val))
        if select_last and self.collector.steps:
            last = str(len(self.collector.steps) - 1)
            self.tree.see(last)
        self._update_code()

    def _update_code(self):
        self.code.delete("1.0", "end")
        self.code.insert("1.0", generate_python(self.collector.steps, self.meta))
        pos = self.code.search("def run(s):", "1.0", "end")      # show the recorded steps, not the boilerplate
        if pos:
            self.code.yview(pos)

    def _sel_indexes(self) -> list[int]:
        return sorted(int(i) for i in self.tree.selection())

    # ---------- sessions / diagnose ----------
    def refresh_sessions(self):
        try:
            self.sessions = list_sessions()
        except Exception as exc:
            self.sessions = []
            self.sess_cb["values"] = []
            self.sess_var.set("")
            self.status.set("SAP not reachable - start SAP Logon and log on, then Refresh. (" + str(exc).splitlines()[0] + ")")
            return
        self.sess_cb["values"] = [s["label"] for s in self.sessions]
        if self.sessions:
            self.sess_cb.current(0)
            self.status.set(f"{len(self.sessions)} SAP session(s) found")
        else:
            self.sess_var.set("")
            self.status.set("SAP Logon is running but no session is logged on")

    def run_diagnose(self):
        s = self._selected_session()
        self.nb.select(1)
        self._log("=== Diagnose ===")
        try:
            self._log(diagnose(s["conn"] if s else None, s["ses"] if s else None))
        except Exception as exc:
            self._log(f"Diagnose failed: {exc}\n{traceback.format_exc()}")

    # ---------- record ----------
    def start_record(self):
        s = self._selected_session()
        if not s:
            self.refresh_sessions()
            s = self._selected_session()
            if not s:
                messagebox.showwarning(APP_TITLE, "No SAP session found.\nStart SAP Logon, log on, then click Refresh.")
                return
        if s.get("disabled"):
            messagebox.showerror(APP_TITLE, "Scripting is disabled by the SAP server for this connection "
                                            "(RZ11 sapgui/user_scripting). Ask Basis to enable it.")
            return
        if self.collector.steps:
            ans = messagebox.askyesnocancel(APP_TITLE, "Keep the existing steps and append the new recording?\n\n"
                                                       "Yes = append    No = start fresh")
            if ans is None:
                return
            if ans is False:
                self.collector.steps.clear()
                self._refresh_tree()
        self.rec = RecorderThread(s["conn"], s["ses"], self.q)
        self.rec.start()
        self.status.set("Starting recorder…")
        self._refresh_buttons()

    def stop_all(self):
        if self.rec:
            self.rec.stop()
            self.status.set("Stopping…")
        if self.proc:
            try:
                self.proc.terminate()
            except Exception:
                pass
            self._log("Playback stopped by user.")

    # ---------- play ----------
    def play(self):
        s = self._selected_session()
        if not s:
            self.refresh_sessions()
            s = self._selected_session()
        if not s:
            messagebox.showwarning(APP_TITLE, "No SAP session found. Log on to SAP, then click Refresh.")
            return
        if not any(not st.get("disabled") for st in self.collector.steps):
            messagebox.showinfo(APP_TITLE, "Nothing to play yet - record some steps first.")
            return
        code = self.code.get("1.0", "end")
        if "DATA = {" in code and re.search(r"^\s+'[^']+': \"\",\s*$", code, re.M):
            if not messagebox.askyesno(APP_TITLE, "Some ${parameters} in DATA are still empty.\n\n"
                                                  "Play anyway?  (No = go to the Python code tab to fill them)"):
                self.nb.select(0)
                return
        path = (self.json_path.with_name(self.json_path.stem + "_replay.py") if self.json_path
                else WORK_DIR / "last_replay.py")
        path.write_text(code, encoding="utf-8")
        write_requirements(path.parent)
        exe = sys.executable
        if exe.lower().endswith("pythonw.exe"):
            cand = Path(exe).with_name("python.exe")
            exe = str(cand) if cand.exists() else exe
        cmd = [exe, "-u", str(path), "--conn", str(s["conn"]), "--ses", str(s["ses"]),
               "--delay", str(self.delay.get()), "--screenshots", str(path.parent / "replay_screenshots")]
        if not self.stop_err.get():
            cmd.append("--keep-going")
        self.nb.select(1)
        self._log(f"=== Play {path.name} on con[{s['conn']}]/ses[{s['ses']}] ===")
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            self.proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                         encoding="utf-8", errors="replace", creationflags=flags,
                                         cwd=str(path.parent))
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not start Python:\n{exc}")
            return
        self.status.set(f"Playing {path} …")
        self._refresh_buttons()
        threading.Thread(target=self._pipe_output, args=(self.proc,), daemon=True).start()

    def _pipe_output(self, proc):
        for line in proc.stdout:
            self.q.put(("play_out", line))
        self.q.put(("play_done", proc.wait()))

    # ---------- queue pump ----------
    def _poll(self):
        changed = False
        try:
            for _ in range(500):
                kind, *rest = self.q.get_nowait()
                if kind == "change":
                    cid, arr, tcode = rest
                    if self.collector.add_event(cid, arr, tcode):
                        changed = True
                elif kind == "meta":
                    self.meta = rest[0]
                elif kind == "recording":
                    self.status.set("● RECORDING - work in SAP GUI, then click Stop")
                    self._log(f"Recording started on {self.meta.get('system')} client {self.meta.get('client')}")
                elif kind == "log":
                    self._log(rest[0])
                elif kind == "error":
                    self._log("ERROR: " + rest[0])
                    self.nb.select(1)
                    messagebox.showerror(APP_TITLE, rest[0].split("\n\n")[0])
                elif kind == "stopped":
                    self.rec = None
                    self._refresh_buttons()
                    self.status.set(f"Recording stopped - {len(self.collector.steps)} step(s). Click Play to replay.")
                    self._log(f"Recording stopped, {len(self.collector.steps)} step(s).")
                elif kind == "play_out":
                    self._log(rest[0])
                elif kind == "play_done":
                    rc = rest[0]
                    self.proc = None
                    self._refresh_buttons()
                    self.status.set("Playback PASSED" if rc == 0 else f"Playback FAILED (exit code {rc}) - see Log")
        except queue.Empty:
            pass
        if changed:
            self._refresh_tree(select_last=True)
            self.status.set(f"● RECORDING - {len(self.collector.steps)} step(s)")
        self.root.after(80, self._poll)

    # ---------- step editing ----------
    def edit_step(self):
        idx = self._sel_indexes()
        if len(idx) != 1:
            return
        s = self.collector.steps[idx[0]]
        if s["type"] == "SP":
            new = simpledialog.askstring(APP_TITLE, f"New value for {s['id']}.{s['name']}\n"
                                                    "(use ${NAME} to make it a parameter)",
                                         initialvalue=str(s["value"]), parent=self.root)
            if new is None:
                return
            old = s["value"]
            if isinstance(old, bool):
                s["value"] = new.strip().lower() in ("1", "true", "yes", "x")
            elif isinstance(old, int) and new.strip().lstrip("-").isdigit():
                s["value"] = int(new)
            else:
                s["value"] = new
        else:
            new = simpledialog.askstring(APP_TITLE, f"Arguments for {s['name']} as a JSON list, e.g. [0] or [2, \"MATNR\"]",
                                         initialvalue=json.dumps(s.get("args", [])), parent=self.root)
            if new is None:
                return
            try:
                args = json.loads(new)
                s["args"] = args if isinstance(args, list) else [args]
            except json.JSONDecodeError as exc:
                messagebox.showerror(APP_TITLE, f"Not valid JSON: {exc}")
                return
        self._refresh_tree()
        self.tree.selection_set(str(idx[0]))

    def delete_steps(self):
        idx = self._sel_indexes()
        for i in reversed(idx):
            del self.collector.steps[i]
        self._refresh_tree()

    def toggle_steps(self):
        idx = self._sel_indexes()
        for i in idx:
            s = self.collector.steps[i]
            s["disabled"] = not s.get("disabled", False)
        self._refresh_tree()
        self.tree.selection_set([str(i) for i in idx])

    def move(self, d):
        idx = self._sel_indexes()
        if len(idx) != 1:
            return
        i, j = idx[0], idx[0] + d
        st = self.collector.steps
        if 0 <= j < len(st):
            st[i], st[j] = st[j], st[i]
            self._refresh_tree()
            self.tree.selection_set(str(j))

    # ---------- files ----------
    def save_json(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", initialdir=str(WORK_DIR),
                                         filetypes=[("Recording", "*.json")])
        if not p:
            return
        doc = {"format": 1, "recorded_at": _dt.datetime.now().isoformat(timespec="seconds"),
               "session": self.meta, "steps": self.collector.steps}
        Path(p).write_text(json.dumps(doc, indent=2, ensure_ascii=False), encoding="utf-8")
        self.json_path = Path(p)
        self.status.set(f"Saved {p}")

    def open_json(self):
        p = filedialog.askopenfilename(initialdir=str(WORK_DIR), filetypes=[("Recording", "*.json")])
        if not p:
            return
        try:
            doc = json.loads(Path(p).read_text(encoding="utf-8"))
            self.collector = StepCollector(list(doc["steps"]))
            self.meta = doc.get("session", {})
        except Exception as exc:
            messagebox.showerror(APP_TITLE, f"Could not open {p}:\n{exc}")
            return
        self.json_path = Path(p)
        self._refresh_tree()
        self.status.set(f"Opened {p} - {len(self.collector.steps)} step(s)")

    def export_py(self):
        p = filedialog.asksaveasfilename(defaultextension=".py", initialdir=str(WORK_DIR),
                                         filetypes=[("Python", "*.py")])
        if p:
            Path(p).write_text(self.code.get("1.0", "end"), encoding="utf-8")
            req = write_requirements(Path(p).parent)
            self.status.set(f"Exported {p}  (+ {req.name})")

    def clear(self):
        if self.collector.steps and not messagebox.askyesno(APP_TITLE, "Remove all steps?"):
            return
        self.collector.steps.clear()
        self.json_path = None
        self._refresh_tree()

    def _on_close(self):
        if self.rec:
            self.rec.stop()
            self.rec.join(timeout=2)
        if self.proc:
            self.proc.terminate()
        self.root.destroy()


def main():
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on high-DPI screens
    except Exception:
        pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
