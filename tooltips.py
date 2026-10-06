"""Hover-popup helper for tkinter widgets (see tooltip_text.py for the text)."""

import re
import tkinter as tk
from typing import Callable

_STATE_RE = re.compile(r"\{\{(\w+)\}\}")


class HoverTip:
    """Show a popup after the mouse rests on a widget.

    resolver(event) returns the text to show for the pointer position, or
    None/"" for none.  For static text use HoverTip.static().
    theme() returns {"bg", "fg", "states": {state: color}}; it is read each
    time a popup opens so dark/light mode changes are picked up.
    """

    DELAY_MS = 450

    def __init__(self, widget, resolver: Callable[[tk.Event], str | None],
                 theme: Callable[[], dict]):
        self._w = widget
        self._resolver = resolver
        self._theme = theme
        self._current: str | None = None
        self._job: str | None = None
        self._win: tk.Toplevel | None = None
        widget.bind("<Motion>", self._on_motion, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    @classmethod
    def static(cls, widget, text: str, theme: Callable[[], dict]) -> "HoverTip":
        return cls(widget, lambda _e: text, theme)

    def _on_motion(self, event) -> None:
        text = self._resolver(event) or None
        if text != self._current:
            self._hide()
            self._current = text
        if text and self._win is None and self._job is None:
            x, y = event.x_root, event.y_root
            self._job = self._w.after(self.DELAY_MS, lambda: self._show(text, x, y))

    def _hide(self, _event=None) -> None:
        if self._job is not None:
            self._w.after_cancel(self._job)
            self._job = None
        if self._win is not None:
            self._win.destroy()
            self._win = None
        if _event is not None:
            self._current = None

    def _show(self, text: str, x: int, y: int) -> None:
        self._job = None
        th = self._theme()
        win = tk.Toplevel(self._w)
        win.wm_overrideredirect(True)
        win.attributes("-topmost", True)

        body = tk.Text(win, wrap="word", width=44, height=1, relief="solid",
                       borderwidth=1, padx=6, pady=4, cursor="arrow",
                       bg=th["bg"], fg=th["fg"], font=("", 9),
                       highlightthickness=0)
        body.tag_configure("title", font=("", 9, "bold"))
        body.pack()

        lines = text.split("\n")
        for i, line in enumerate(lines):
            pos = 0
            for m in _STATE_RE.finditer(line):
                body.insert("end", line[pos:m.start()])
                state = m.group(1)
                tag = "st_" + state
                color = th["states"].get(state)
                if color:
                    body.tag_configure(tag, background=color, foreground=th["fg"],
                                       font=("", 9, "bold"))
                body.insert("end", f" {state} ", tag if color else ())
                pos = m.end()
            body.insert("end", line[pos:], "title" if i == 0 and len(lines) > 1
                        and not _STATE_RE.search(line) else ())
            if i < len(lines) - 1:
                body.insert("end", "\n")

        # Size to content: wrapped display lines
        width = int(body.cget("width"))
        n = sum(max(1, -(-len(_STATE_RE.sub(lambda m: f" {m.group(1)} ", ln)) // (width - 4)))
                for ln in lines)
        body.configure(height=n, state="disabled")

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        sx, sy = win.winfo_screenwidth(), win.winfo_screenheight()
        px = min(x + 14, sx - w - 4)
        py = y + 18 if y + 18 + h < sy else y - h - 8
        win.wm_geometry(f"+{max(0, px)}+{max(0, py)}")
        self._win = win
