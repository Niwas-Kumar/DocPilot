# -*- coding: utf-8 -*-
"""Desktop entry point for the packaged DocPilot app.

Packaging only: it starts the very same FastAPI app the server does, picks a
free port, and shows a small status window. Nothing about the pipelines
changes.

The window replaces the old console. That is not only cosmetic - a console
build makes a black terminal the app's face, and closing it is the only
"stop" affordance it offers. It also makes the build windowed, so:

  * sys.stdout / sys.stderr are None under PyInstaller noconsole mode, and any
    stray print() in the backend would raise AttributeError. _Ring below
    replaces them with a capped buffer, which doubles as the text shown if the
    server dies before it can bind.
  * a crash has nowhere to print, so the window shows it instead.

Run with --console (or DOCPILOT_CONSOLE=1) to get the old terminal behaviour
back for troubleshooting.
"""
import ctypes
import os
import socket
import sys
import threading
import time
import traceback
import webbrowser

VERSION = "2.0-beta"
DEFAULT_PORT = 8930

BG = "#0e1116"
FG = "#ffffff"
MUTED = "#9db2c8"
FAINT = "#8296ac"
GREEN = "#3ddc84"
AMBER = "#f2b807"
AMBER_HOT = "#ffc720"
RED = "#ff6b6b"
CHIP_BG = "#1b2129"
CHIP_FG = "#c9d6e4"
DISABLED = "#333a44"

W, H = 660, 420


class _Ring:
    """A stdout stand-in that keeps only the last few KB.

    Windowed builds have no console; without this every print() in the backend
    raises, and a startup failure leaves nothing to show the user.
    """

    def __init__(self, cap=8000):
        self._buf = []
        self._cap = cap

    def write(self, text):
        self._buf.append(text)
        total = sum(len(t) for t in self._buf)
        while total > self._cap and len(self._buf) > 1:
            total -= len(self._buf.pop(0))
        return len(text)

    def flush(self):
        pass

    def isatty(self):
        return False

    def tail(self, lines=6):
        return "".join(self._buf).strip().splitlines()[-lines:]


def _install_stdio():
    ring = _Ring()
    if sys.stdout is None:
        sys.stdout = ring
    if sys.stderr is None:
        sys.stderr = ring
    return ring


def _app_root():
    # frozen builds keep the whole tree under sys._MEIPASS; every module in
    # this project derives its asset paths from __file__, which PyInstaller
    # maps into that same tree, so nothing needs redirecting
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def _free_port(preferred=DEFAULT_PORT):
    """The preferred port if it is free, else the next one that is."""
    for port in range(preferred, preferred + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def _port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.4)
        return s.connect_ex(("127.0.0.1", port)) == 0


def _serve(port, on_error):
    """Run uvicorn. Any failure is handed back to the UI thread."""
    try:
        import uvicorn
        from backend.app import app
        uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")
    except BaseException:                        # noqa: BLE001 - shown to user
        traceback.print_exc()
        on_error(traceback.format_exc(limit=3))


APP_MUTEX = "DocPilot_PW_OnlyIAS_Running"
_mutex_handle = None


def _claim_mutex():
    """Publish a named mutex the installer can see.

    The installer declares the same name as AppMutex, so Setup can say "close
    DocPilot first" up front. Without it Setup falls back to Restart Manager,
    which asks an app to shut itself down gracefully - a request a Python/Tk
    process does not answer, so Setup reports it "was unable to automatically
    close all applications" and dumps the choice on the user.

    The handle is kept in a module global on purpose: let it be collected and
    Windows drops the mutex while the app is still running.
    """
    global _mutex_handle
    try:
        _mutex_handle = ctypes.windll.kernel32.CreateMutexW(None, False,
                                                            APP_MUTEX)
    except Exception:
        _mutex_handle = None


def _enable_dpi_awareness():
    """Tell Windows we paint real pixels, and report the scale factor.

    Without this the shell bitmap-stretches the whole window on any display
    above 100%, which is exactly the soft, fuzzy look a bundled tool is judged
    for. Being aware means nothing is scaled for us, so every coordinate and
    font size below is multiplied by the returned factor instead.
    """
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)      # per-monitor v1
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            return 1.0
    for getter in ("GetDpiForSystem",):
        try:
            return max(1.0, getattr(ctypes.windll.user32, getter)() / 96.0)
        except Exception:
            pass
    try:
        dc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)     # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, dc)
        return max(1.0, dpi / 96.0)
    except Exception:
        return 1.0


def _dark_titlebar(win):
    """Windows 11 dark title bar. Silently skipped on older builds."""
    try:
        win.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        flag = ctypes.c_int(1)
        for attr in (20, 19):                    # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, attr, ctypes.byref(flag), ctypes.sizeof(flag)) == 0:
                return
    except Exception:
        pass


def _round_rect(canvas, x1, y1, x2, y2, r, **kw):
    """Tk has no rounded rectangle; a smoothed polygon is the usual stand-in."""
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(pts, smooth=True, **kw)


class Window:
    def __init__(self, port, ring, scale=1.0):
        import tkinter as tk
        from tkinter import font as tkfont

        self.s = scale
        self.port = port
        self.ring = ring
        self.url = "http://127.0.0.1:%d/" % port
        self.state = "starting"
        self.detail = ""
        self.opened = False
        self._btn_id = None
        self._btn_box = (0, 0, 0, 0)

        self.root = tk.Tk()
        self.root.title("DocPilot")
        self.root.configure(bg=BG)
        self.root.resizable(False, False)
        self.pw, self.ph = int(W * scale), int(H * scale)
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry("%dx%d+%d+%d" % (
            self.pw, self.ph, max(0, (sw - self.pw) // 2),
            max(0, (sh - self.ph) // 2 - 40)))

        ico = os.path.join(_app_root(), "assets", "icon", "docpilot.ico")
        if os.path.isfile(ico):
            try:
                self.root.iconbitmap(ico)
            except Exception:
                pass

        # negative size = pixels, so the layout stays predictable instead of
        # depending on Tk's own point-to-pixel guess
        def px(n):
            return tkfont.Font(family="Segoe UI", size=-int(round(n * scale)))

        self.f_title = px(26)
        self.f_title.configure(weight="bold")
        self.f_body = px(15)
        self.f_btn = px(17)
        self.f_btn.configure(weight="bold")
        self.f_small = px(13)
        self.f_mono = tkfont.Font(family="Consolas",
                                  size=-int(round(14 * scale)))

        self.c = tk.Canvas(self.root, width=self.pw, height=self.ph, bg=BG,
                           highlightthickness=0, bd=0)
        self.c.pack(fill="both", expand=True)
        self.c.bind("<Button-1>", self._click)
        self.c.bind("<Motion>", self._hover)

        _dark_titlebar(self.root)
        self.root.protocol("WM_DELETE_WINDOW", self.quit)
        self.draw()

    # ---------------------------------------------------------------- paint
    def draw(self):
        c = self.c
        c.delete("all")
        S = self.s
        cx = self.pw / 2

        def y(v):
            return v * S

        running = self.state == "running"
        failed = self.state == "failed"
        dot = GREEN if running else (RED if failed else AMBER)
        word = "running" if running else ("failed" if failed else "starting")

        # title + status, measured so the pair sits centred as one group
        tw = self.f_title.measure("DocPilot")
        sw_ = self.f_title.measure(word)
        gap, r = y(16), y(6)
        x = cx - (tw + gap + 2 * r + y(10) + sw_) / 2
        c.create_text(x, y(124), text="DocPilot", anchor="w", fill=FG,
                      font=self.f_title)
        dx = x + tw + gap + r
        c.create_oval(dx - r, y(124) - r, dx + r, y(124) + r, fill=dot,
                      outline="")
        c.create_text(dx + r + y(10), y(124), text=word, anchor="w", fill=dot,
                      font=self.f_title)

        # address line
        if failed:
            c.create_text(cx, y(170), text="DocPilot could not start.",
                          fill=MUTED, font=self.f_body)
        elif running:
            prefix = "The app is running locally at"
            pw = self.f_body.measure(prefix)
            uw = self.f_mono.measure(self.url)
            pad = y(11)
            gx = cx - (pw + y(12) + uw + 2 * pad) / 2
            c.create_text(gx, y(170), text=prefix, anchor="w", fill=MUTED,
                          font=self.f_body)
            bx = gx + pw + y(12)
            _round_rect(c, bx, y(157), bx + uw + 2 * pad, y(184), y(6),
                        fill=CHIP_BG, outline="")
            c.create_text(bx + pad, y(170), text=self.url, anchor="w",
                          fill=CHIP_FG, font=self.f_mono)
        else:
            c.create_text(cx, y(170), text="Starting the local server...",
                          fill=MUTED, font=self.f_body)

        # button
        bw, bh = y(214), y(54)
        x1, y1 = cx - bw / 2, y(212)
        self._btn_box = (x1, y1, x1 + bw, y1 + bh)
        self._btn_id = _round_rect(c, x1, y1, x1 + bw, y1 + bh, y(8),
                                   fill=AMBER if running else DISABLED,
                                   outline="")
        c.create_text(cx, y1 + bh / 2, text="Open in browser",
                      fill="#1a1a1a" if running else "#6b7480", font=self.f_btn)

        # footer
        if failed:
            for i, line in enumerate(self.detail.splitlines()[-5:]):
                c.create_text(cx, y(300 + i * 16), text=line[:96], fill=FAINT,
                              font=self.f_mono)
        else:
            c.create_text(cx, y(318),
                          text="Keep this window open while you use the app.",
                          fill=FAINT, font=self.f_small)
            c.create_text(cx, y(344), text="Closing it stops the server.",
                          fill=FAINT, font=self.f_small)

        c.create_text(self.pw - y(14), self.ph - y(12), text="v" + VERSION,
                      anchor="e", fill="#4a5563", font=self.f_small)

    # --------------------------------------------------------------- events
    def _in_button(self, x, y):
        x1, y1, x2, y2 = self._btn_box
        return x1 <= x <= x2 and y1 <= y <= y2

    def _click(self, ev):
        if self.state == "running" and self._in_button(ev.x, ev.y):
            webbrowser.open(self.url)

    def _hover(self, ev):
        if self.state != "running":
            self.c.configure(cursor="")
            return
        over = self._in_button(ev.x, ev.y)
        self.c.configure(cursor="hand2" if over else "")
        if self._btn_id:
            self.c.itemconfigure(self._btn_id, fill=AMBER_HOT if over else AMBER)

    def quit(self):
        # the server runs in a daemon thread; a hard exit is what "closing it
        # stops the server" has to mean
        try:
            self.root.destroy()
        finally:
            os._exit(0)

    # ---------------------------------------------------------------- state
    def fail(self, detail):
        self.state, self.detail = "failed", detail
        try:
            self.root.after(0, self.draw)
        except Exception:
            pass

    def _poll(self, waited):
        if self.state == "failed":
            return
        if _port_open(self.port):
            self.state = "running"
            self.draw()
            if not self.opened:
                self.opened = True
                webbrowser.open(self.url)
            return
        if waited > 90:
            self.fail("\n".join(self.ring.tail()) or
                      "The server did not start within 90 seconds.")
            return
        self.root.after(400, lambda: self._poll(waited + 0.4))

    def run(self):
        self.root.after(300, lambda: self._poll(0.0))
        self.root.mainloop()


def _console_main(port):
    print("=" * 62)
    print("  DocPilot %s  -  PW OnlyIAS Document Studio" % VERSION)
    print("=" * 62)
    print("  Opening http://127.0.0.1:%d" % port)
    print("  Keep this window open while you work.")
    print("  Close it to stop DocPilot.")
    print("=" * 62)
    threading.Thread(
        target=lambda: (time.sleep(2),
                        webbrowser.open("http://127.0.0.1:%d" % port)),
        daemon=True).start()
    _serve(port, lambda detail: None)


def main():
    ring = _install_stdio()
    _claim_mutex()

    root = _app_root()
    os.chdir(root)
    if root not in sys.path:
        sys.path.insert(0, root)

    port = _free_port()

    if "--console" in sys.argv or os.environ.get("DOCPILOT_CONSOLE") == "1":
        _console_main(port)
        return

    win = Window(port, ring, scale=_enable_dpi_awareness())
    threading.Thread(target=_serve, args=(port, win.fail), daemon=True).start()
    win.run()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        traceback.print_exc()
        try:
            import tkinter.messagebox as mb
            mb.showerror("DocPilot", "DocPilot could not start:\n\n%s" % exc)
        except Exception:
            pass
