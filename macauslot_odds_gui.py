#!/usr/bin/env python3
"""GUI launcher for the Macauslot odds scraper."""

from __future__ import annotations

import argparse
import contextlib
import datetime as dt
import os
import queue
import random
import sys
import threading
import traceback
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, BooleanVar, Button, Checkbutton, Entry, Frame, Label, StringVar, Tk, messagebox
from tkinter.scrolledtext import ScrolledText

import scrape_macauslot_odds


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        executable = Path(sys.executable).resolve()
        for parent in executable.parents:
            if parent.name.endswith(".app"):
                return parent.parent
        return executable.parent
    return Path(__file__).resolve().parent


ROOT = app_base_dir()
OUTPUT_FILE = ROOT / "outputs" / "清洗后的表格.xlsx"
AFTER_SALES_CONTACT = "19236521940"

COLOR_BG = "#F3F6FA"
COLOR_SURFACE = "#FFFFFF"
COLOR_HEADER = "#0F172A"
COLOR_TEXT = "#111827"
COLOR_MUTED = "#64748B"
COLOR_BORDER = "#D9E2EC"
COLOR_ACCENT = "#16A34A"
COLOR_ACCENT_DARK = "#15803D"
COLOR_BLUE = "#2563EB"
COLOR_BLUE_DARK = "#1D4ED8"
COLOR_DANGER = "#DC2626"
COLOR_DANGER_DARK = "#B91C1C"
COLOR_LOG_BG = "#0B1220"
COLOR_LOG_TEXT = "#E5E7EB"

FONT_BASE = ("Microsoft YaHei UI", 10)
FONT_TITLE = ("Microsoft YaHei UI", 18, "bold")
FONT_SECTION = ("Microsoft YaHei UI", 11, "bold")
FONT_MONO = ("Consolas", 10)


def normalize_interval_minutes(value: object) -> int:
    try:
        return max(1, int(str(value).strip()))
    except Exception:
        return 1


def normalize_random_delay_seconds(value: object) -> int:
    try:
        return max(0, int(str(value).strip()))
    except Exception:
        return 30


def calculate_next_delay_ms(interval_minutes: int, random_delay_seconds: int, randint=random.randint) -> int:
    base_seconds = normalize_interval_minutes(interval_minutes) * 60
    extra_seconds = randint(0, normalize_random_delay_seconds(random_delay_seconds))
    return (base_seconds + extra_seconds) * 1000


def should_schedule_next_timer_run(completed_runs: int) -> bool:
    return True


def format_status_summary(
    running: bool,
    timer_active: bool,
    timer_run_count: int = 0,
    proxy_host: str = "",
    proxy_port: str = "",
    next_run_text: str = "",
) -> str:
    state = "运行中" if running else "定时中" if timer_active else "空闲"
    parts = [state]
    if timer_active or timer_run_count:
        parts.append(f"已执行 {timer_run_count} 次")
    if next_run_text:
        parts.append(f"下次 {next_run_text}")
    host = proxy_host.strip()
    port = proxy_port.strip()
    parts.append(f"代理 {host}:{port}" if host and port else "直连")
    return "  |  ".join(parts)


def build_scraper_args(
    output_file: Path,
    proxy_host: str = "",
    proxy_port: str = "",
    no_verify_ssl: bool = False,
) -> argparse.Namespace:
    return argparse.Namespace(
        url=scrape_macauslot_odds.ENTRY_URL,
        lang="cn",
        output=str(output_file),
        format="visible",
        types=None,
        proxy_host=proxy_host.strip(),
        proxy_port=proxy_port.strip(),
        no_verify_ssl=bool(no_verify_ssl),
    )


def is_startup_blocked() -> bool:
    try:
        return not scrape_macauslot_odds.is_run_allowed()
    except RuntimeError:
        # Let the user open the GUI and configure proxy settings; run_scraper
        # will report the network-time error before scraping starts.
        return False


def should_play_success_sound(scheduled: bool, result: scrape_macauslot_odds.RunScrapeResult | None) -> bool:
    return bool(scheduled and result and result.changed)


def after_sales_contact_text() -> str:
    return f"\u552e\u540e\u8054\u7cfb\u65b9\u5f0f\uff1a{AFTER_SALES_CONTACT}"


def play_success_sound() -> None:
    if os.name != "nt":
        return
    try:
        import winsound

        winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        return


class QueueWriter:
    def __init__(self, log_queue: queue.Queue[str]) -> None:
        self.log_queue = log_queue

    def write(self, text: str) -> None:
        if text:
            self.log_queue.put(text)

    def flush(self) -> None:
        pass


class App:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title("澳彩赔率更新控制台")
        self.root.geometry("980x680")
        self.root.minsize(900, 620)
        self.root.configure(bg=COLOR_BG)
        self.root.option_add("*Font", FONT_BASE)
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.timer_active = False
        self.startup_blocked = is_startup_blocked()
        self.timer_after_id: str | None = None
        self.next_run_text = ""
        self.timer_run_count = 0
        self.interval_var = StringVar(value="3")
        self.random_delay_var = StringVar(value="30")
        self.proxy_host_var = StringVar(value="")
        self.proxy_port_var = StringVar(value="")
        self.no_verify_ssl_var = BooleanVar(value=False)

        header = Frame(self.root, bg=COLOR_HEADER)
        header.pack(fill="x")
        header.columnconfigure(0, weight=1)

        title_group = Frame(header, bg=COLOR_HEADER)
        title_group.grid(row=0, column=0, sticky="ew", padx=20, pady=18)
        Label(
            title_group,
            text="澳彩赔率更新控制台",
            bg=COLOR_HEADER,
            fg="#F8FAFC",
            font=FONT_TITLE,
            anchor="w",
        ).pack(fill="x")
        initial_status = "点击“立即更新”抓取全場三合一賠率"
        self.status = Label(
            title_group,
            text=initial_status,
            bg=COLOR_HEADER,
            fg="#CBD5E1",
            anchor="w",
        )
        self.status.pack(fill="x", pady=(4, 0))
        self.contact_label = Label(
            title_group,
            text=after_sales_contact_text(),
            bg=COLOR_HEADER,
            fg="#FDE68A",
            font=("Microsoft YaHei UI", 10, "bold"),
            anchor="w",
        )
        self.contact_label.pack(fill="x", pady=(6, 0))

        self.summary_label = Label(
            header,
            text="",
            bg="#111827",
            fg="#BBF7D0",
            font=("Microsoft YaHei UI", 10, "bold"),
            padx=14,
            pady=8,
        )
        self.summary_label.grid(row=0, column=1, sticky="e", padx=20, pady=18)

        body = Frame(self.root, bg=COLOR_BG)
        body.pack(fill=BOTH, expand=True, padx=16, pady=16)

        cards = Frame(body, bg=COLOR_BG)
        cards.pack(fill="x")
        for col in range(3):
            cards.columnconfigure(col, weight=1)

        control_card = self.create_card(cards, "执行控制")
        control_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.start_button = self.create_button(control_card, "立即更新", self.start, COLOR_ACCENT, COLOR_ACCENT_DARK)
        self.start_button.pack(fill="x", pady=(2, 8))
        self.timer_start_button = self.create_button(control_card, "开始定时", self.start_timer, COLOR_BLUE, COLOR_BLUE_DARK)
        self.timer_start_button.pack(fill="x", pady=(0, 8))
        self.timer_stop_button = self.create_button(control_card, "停止定时", self.stop_timer, COLOR_DANGER, COLOR_DANGER_DARK)
        self.timer_stop_button.pack(fill="x")

        schedule_card = self.create_card(cards, "定时策略")
        schedule_card.grid(row=0, column=1, sticky="nsew", padx=10)
        self.create_field(schedule_card, "定时间隔（分钟）", self.interval_var, "最低 1 分钟")
        self.create_field(schedule_card, "随机延迟（秒）", self.random_delay_var, "每轮额外等待 0 到该秒数")

        network_card = self.create_card(cards, "网络代理")
        network_card.grid(row=0, column=2, sticky="nsew", padx=(10, 0))
        self.create_field(network_card, "代理 IP / 地址", self.proxy_host_var, "留空表示直连")
        self.create_field(network_card, "端口号", self.proxy_port_var, "例如 7890")
        Checkbutton(
            network_card,
            text="忽略 SSL 证书校验",
            variable=self.no_verify_ssl_var,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            activebackground=COLOR_SURFACE,
            activeforeground=COLOR_TEXT,
            selectcolor=COLOR_SURFACE,
            anchor="w",
            cursor="hand2",
        ).pack(fill="x", pady=(0, 4))
        Label(
            network_card,
            text="仅在可信代理使用自签证书时开启",
            bg=COLOR_SURFACE,
            fg=COLOR_MUTED,
            anchor="w",
        ).pack(fill="x")

        log_card = self.create_card(body, "运行日志")
        log_card.pack(fill=BOTH, expand=True, pady=(16, 0))
        self.log = ScrolledText(
            log_card,
            height=24,
            wrap="word",
            bg=COLOR_LOG_BG,
            fg=COLOR_LOG_TEXT,
            insertbackground=COLOR_LOG_TEXT,
            relief="flat",
            borderwidth=0,
            font=FONT_MONO,
        )
        self.log.pack(fill=BOTH, expand=True)

        footer = Frame(self.root, bg=COLOR_BG)
        footer.pack(fill="x", padx=16, pady=(0, 14))
        Label(
            footer,
            text=f"输出文件：{OUTPUT_FILE}",
            anchor="w",
            bg=COLOR_BG,
            fg=COLOR_MUTED,
        ).pack(side=LEFT, fill="x", expand=True)
        self.refresh_status_summary()
        self.root.after(100, self.drain_log_queue)

    def create_card(self, parent, title: str) -> Frame:
        card = Frame(
            parent,
            bg=COLOR_SURFACE,
            highlightbackground=COLOR_BORDER,
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        Label(
            card,
            text=title,
            bg=COLOR_SURFACE,
            fg=COLOR_TEXT,
            font=FONT_SECTION,
            anchor="w",
        ).pack(fill="x", pady=(0, 10))
        return card

    def create_button(self, parent, text: str, command, bg: str, active_bg: str) -> Button:
        return Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg="#FFFFFF",
            activebackground=active_bg,
            activeforeground="#FFFFFF",
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=9,
            cursor="hand2",
            font=("Microsoft YaHei UI", 10, "bold"),
        )

    def create_field(self, parent, label: str, variable: StringVar, hint: str) -> None:
        field = Frame(parent, bg=COLOR_SURFACE)
        field.pack(fill="x", pady=(0, 12))
        Label(field, text=label, bg=COLOR_SURFACE, fg=COLOR_TEXT, anchor="w").pack(fill="x")
        Entry(
            field,
            textvariable=variable,
            bg="#FFFFFF",
            fg=COLOR_TEXT,
            relief="solid",
            borderwidth=1,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            highlightcolor=COLOR_BLUE,
        ).pack(fill="x", pady=(5, 4), ipady=4)
        Label(field, text=hint, bg=COLOR_SURFACE, fg=COLOR_MUTED, anchor="w").pack(fill="x")

    def refresh_status_summary(self) -> None:
        self.summary_label.configure(
            text=format_status_summary(
                running=self.running,
                timer_active=self.timer_active,
                timer_run_count=self.timer_run_count,
                proxy_host=self.proxy_host_var.get(),
                proxy_port=self.proxy_port_var.get(),
                next_run_text=self.next_run_text,
            )
        )

    def start(self) -> None:
        self.start_run(scheduled=False, clear_log=True)

    def start_timer(self) -> None:
        self.timer_active = True
        self.next_run_text = ""
        self.timer_run_count = 0
        self.log.delete("1.0", END)
        minutes = normalize_interval_minutes(self.interval_var.get())
        self.interval_var.set(str(minutes))
        random_delay = normalize_random_delay_seconds(self.random_delay_var.get())
        self.random_delay_var.set(str(random_delay))
        self.log_queue.put(f"定时更新已启动，间隔 {minutes} 分钟，随机延迟 0-{random_delay} 秒。\n")
        self.refresh_status_summary()
        self.start_run(scheduled=True, clear_log=False)

    def stop_timer(self) -> None:
        self.timer_active = False
        self.next_run_text = ""
        self.timer_run_count = 0
        if self.timer_after_id:
            self.root.after_cancel(self.timer_after_id)
            self.timer_after_id = None
        if self.running:
            self.status.configure(text="本轮执行完成后停止定时")
        else:
            self.status.configure(text="定时已停止")
        self.refresh_status_summary()
        self.update_buttons()

    def start_run(self, scheduled: bool, clear_log: bool) -> None:
        if self.running:
            self.log_queue.put("已有更新任务正在执行，跳过本次启动。\n")
            return
        if clear_log:
            self.log.delete("1.0", END)
        self.running = True
        self.next_run_text = ""
        self.status.configure(text="正在更新，请稍候...")
        self.refresh_status_summary()
        self.update_buttons()
        args = build_scraper_args(
            OUTPUT_FILE,
            proxy_host=self.proxy_host_var.get(),
            proxy_port=self.proxy_port_var.get(),
            no_verify_ssl=self.no_verify_ssl_var.get(),
        )
        thread = threading.Thread(target=self.run_scraper, args=(args, scheduled), daemon=True)
        thread.start()

    def run_scraper(self, args: argparse.Namespace, scheduled: bool) -> None:
        writer = QueueWriter(self.log_queue)
        success = False
        result = None
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                result = scrape_macauslot_odds.run_scraper(args)
            success = True
            self.log_queue.put("\n更新完成。\n")
            if should_play_success_sound(scheduled, result):
                play_success_sound()
        except Exception:
            self.log_queue.put("\n执行失败：\n")
            self.log_queue.put(traceback.format_exc())
        finally:
            self.root.after(0, lambda: self.finish_run(scheduled, success))

    def finish_run(self, scheduled: bool, success: bool) -> None:
        self.running = False
        if scheduled:
            self.timer_run_count += 1
        if scheduled and self.timer_active:
            minutes = normalize_interval_minutes(self.interval_var.get())
            self.interval_var.set(str(minutes))
            random_delay = normalize_random_delay_seconds(self.random_delay_var.get())
            self.random_delay_var.set(str(random_delay))
            delay_ms = calculate_next_delay_ms(minutes, random_delay)
            next_time = dt.datetime.now() + dt.timedelta(milliseconds=delay_ms)
            self.next_run_text = f"{next_time:%H:%M:%S}"
            self.status.configure(text=f"定时运行中，下次执行：{next_time:%H:%M:%S}")
            self.log_queue.put(f"下次执行时间：{next_time:%Y-%m-%d %H:%M:%S}\n")
            self.timer_after_id = self.root.after(delay_ms, self.run_scheduled)
        else:
            self.next_run_text = ""
            self.status.configure(text="更新完成" if success else "执行失败")
            if success and not scheduled:
                messagebox.showinfo("完成", f"Excel 已更新：\n{OUTPUT_FILE}")
            elif not scheduled:
                messagebox.showerror("失败", "执行失败，请查看日志。")
        self.refresh_status_summary()
        self.update_buttons()

    def run_scheduled(self) -> None:
        self.timer_after_id = None
        if not self.timer_active:
            return
        self.start_run(scheduled=True, clear_log=False)

    def update_buttons(self) -> None:
        self.start_button.configure(state="disabled" if self.startup_blocked or self.running or self.timer_active else "normal")
        self.timer_start_button.configure(state="disabled" if self.startup_blocked or self.timer_active or self.running else "normal")
        self.timer_stop_button.configure(state="normal" if self.timer_active else "disabled")

    def drain_log_queue(self) -> None:
        while True:
            try:
                text = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log.insert(END, text)
            self.log.see(END)
        self.root.after(100, self.drain_log_queue)

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    App().run()
