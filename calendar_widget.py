"""
Simple pop-up calendar date-picker for tkinter.

Usage:
    CalendarPicker(parent_widget, callback, initial_date=None)

    callback is called with the selected date as a "YYYY-MM-DD" string.
    initial_date is a datetime.date object (defaults to today).
"""

import calendar
import tkinter as tk
from tkinter import ttk
from datetime import date


class CalendarPicker(tk.Toplevel):
    """Modal date-picker window."""

    def __init__(self, parent, callback, initial_date: date | None = None):
        super().__init__(parent)
        self.title("Select Date")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._callback = callback
        self._today    = date.today()
        d = initial_date or self._today
        self._year  = d.year
        self._month = d.month

        self._frame = ttk.Frame(self, padding=8)
        self._frame.pack()
        self._render()

        # Centre over parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, px)}+{max(0, py)}")

    # ── Rendering ─────────────────────────────────────────────────────────────

    def _render(self) -> None:
        for w in self._frame.winfo_children():
            w.destroy()

        # ── Navigation header ─────────────────────────────────────────────────
        hdr = ttk.Frame(self._frame)
        hdr.grid(row=0, column=0, columnspan=7, sticky="ew", pady=(0, 6))

        ttk.Button(hdr, text="◀", width=3,
                   command=self._prev_month).pack(side="left")
        ttk.Label(
            hdr,
            text=f"{calendar.month_name[self._month]}  {self._year}",
            font=("", 10, "bold"), anchor="center",
        ).pack(side="left", expand=True, fill="x", padx=8)
        ttk.Button(hdr, text="▶", width=3,
                   command=self._next_month).pack(side="right")

        # ── Day-of-week labels ────────────────────────────────────────────────
        for col, abbr in enumerate(["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]):
            ttk.Label(self._frame, text=abbr, width=4, anchor="center",
                      foreground="#555").grid(row=1, column=col, padx=1)

        # ── Day buttons ───────────────────────────────────────────────────────
        weeks = calendar.monthcalendar(self._year, self._month)
        for r, week in enumerate(weeks, start=2):
            for c, day in enumerate(week):
                if day == 0:
                    ttk.Label(self._frame, width=4).grid(
                        row=r, column=c, padx=1, pady=1)
                    continue
                is_today = (
                    day        == self._today.day   and
                    self._month == self._today.month and
                    self._year  == self._today.year
                )
                btn = tk.Button(
                    self._frame, text=str(day), width=3, relief="flat",
                    bg="#1565C0" if is_today else "#F5F5F5",
                    fg="white"  if is_today else "black",
                    activebackground="#BBDEFB",
                    command=lambda d=day: self._select(d),
                    cursor="hand2",
                )
                btn.grid(row=r, column=c, padx=1, pady=1)

        # ── "Today" shortcut ──────────────────────────────────────────────────
        ttk.Button(
            self._frame, text="Today", command=self._go_today,
        ).grid(row=len(weeks) + 2, column=0, columnspan=7,
               pady=(8, 0), sticky="ew")

    # ── Actions ───────────────────────────────────────────────────────────────

    def _select(self, day: int) -> None:
        selected = date(self._year, self._month, day)
        self._callback(selected.strftime("%Y-%m-%d"))
        self.destroy()

    def _go_today(self) -> None:
        self._year  = self._today.year
        self._month = self._today.month
        self._render()

    def _prev_month(self) -> None:
        if self._month == 1:
            self._month = 12
            self._year -= 1
        else:
            self._month -= 1
        self._render()

    def _next_month(self) -> None:
        if self._month == 12:
            self._month = 1
            self._year += 1
        else:
            self._month += 1
        self._render()
