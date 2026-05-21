#!/usr/bin/env python3
"""Scrape Macauslot soccer odds data and export it to Excel.

The public page renders its tables from JSON feeds. This script fetches those
feeds directly, including the lightweight Aliyun WAF cookie challenge used by
the site when it returns a browser verification page.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import gzip
import json
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from http.cookiejar import Cookie, CookieJar
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_URL = "https://www.macauslot.com"
ENTRY_URL = f"{BASE_URL}/cn/soccer/odds_in.html"
REALTIME_URL = f"{BASE_URL}/soccer/json/realtime"


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        executable = Path(sys.executable).resolve()
        for parent in executable.parents:
            if parent.name.endswith(".app"):
                return parent.parent
        return executable.parent
    return Path(__file__).resolve().parent


def resource_path(relative_path: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).resolve().parent / relative_path


SAMPLE_TEMPLATE_PATH = resource_path("assets/sample_template.xlsx")
RUN_BLOCK_DATETIME = dt.datetime(2026, 5, 18, 0, 0, 0)
# Network time checking is disabled for the no-time-limit build.
# NETWORK_TIME_URLS = (
#     "http://www.baidu.com/",
#     "http://www.qq.com/",
#     "http://www.microsoft.com/",
# )
# NETWORK_TIME_TIMEOUT_SECONDS = 5
WAF_XOR_KEY = "3000176000856006061501533003690027800375"
WAF_BOX = [
    0x0F,
    0x23,
    0x1D,
    0x18,
    0x21,
    0x10,
    0x01,
    0x26,
    0x0A,
    0x09,
    0x13,
    0x1F,
    0x28,
    0x1B,
    0x16,
    0x17,
    0x19,
    0x0D,
    0x06,
    0x0B,
    0x27,
    0x12,
    0x14,
    0x08,
    0x0E,
    0x15,
    0x20,
    0x1A,
    0x02,
    0x1E,
    0x07,
    0x04,
    0x11,
    0x05,
    0x03,
    0x1C,
    0x22,
    0x25,
    0x0C,
    0x24,
]


ODDS_TYPE_NAMES = {
    "index": "全場三合一賠率",
    "half_index": "上半場三合一賠率",
    "half_c1st-2in1": "上半場角球數賠率",
    "half_correctscore": "上半場波膽",
    "half_correctscorespecial": "上半場波膽組合",
    "half_oddeven": "上半場入球單/雙數",
    "half_numberofgoals": "上半場球隊入球數",
    "2in1_corn": "角球數賠率",
    "correctscore": "波膽",
    "correctscorespecial": "波膽組合",
    "halffull": "上半場/全場賽果",
    "oddeven": "入球單/雙數",
    "totalgoals": "全場入球總數",
    "halffulltotalscore": "上/下半場入球較多",
    "numberofgoals": "球隊入球數",
    "firstgoalteam": "最先入球球隊",
    "firstscorer": "首名入球球員",
    "EX_Penalty": "加時/十二碼",
    "bir": "即場投注",
    "FUNC_GJPL": "總冠軍球隊",
    "FUNC_XZGJ": "小組冠軍",
    "FUNC_XZSCM": "小組首次名",
    "FUNC_JXJ": "金靴獎",
    "FUNC_BB_JJQD": "晉級球隊",
    "FUNC_JSMB": "決賽孖寶",
}

BASE_ODDS_TYPES = [
    "index",
    "half_index",
    "half_c1st-2in1",
    "half_correctscore",
    "half_correctscorespecial",
    "half_oddeven",
    "half_numberofgoals",
    "2in1_corn",
    "correctscore",
    "correctscorespecial",
    "halffull",
    "oddeven",
    "totalgoals",
    "halffulltotalscore",
    "numberofgoals",
    "firstgoalteam",
    "firstscorer",
    "EX_Penalty",
    "bir",
]

JSON_CODE_MAP = {
    "half_c1st-2in1": "c1st-2in1",
    "half_correctscore": "cs",
    "half_correctscorespecial": "ss",
    "half_oddeven": "oe",
    "half_numberofgoals": "tt",
    "2in1_corn": "c-2in1",
    "correctscore": "CS",
    "correctscorespecial": "SS",
    "halffull": "HF",
    "oddeven": "OE",
    "numberofgoals": "TT",
    "totalgoals": "TG",
    "halffulltotalscore": "HG",
    "firstgoalteam": "SF",
    "firstscorer": "FS",
    "EX_Penalty": "EX_Penalty",
}

SPECIAL_MENU_IDS = {"GJPL", "XZGJ", "XZSCM", "JXJ", "BB_JJQD", "JSMB"}

SUMMARY_COLUMNS = [
    "工作表",
    "玩法名称",
    "事件数",
    "赔率行数",
    "状态",
    "备注",
]

VISIBLE_COLUMNS = [
    "类别",
    "开赛时间",
    "主/客",
    "市场",
    "选项",
    "盘口",
    "赔率",
]

SAMPLE_HEADERS = {
    1: "类别",
    2: "比赛时间",
    3: "主/客",
    4: "让球盘",
    10: "录入时间",
    11: "上/下盘",
    15: "录入时间",
    16: "标准盘",
    17: "录入时间",
    18: "比分",
}
SAMPLE_SHEET_NAME = "Sheet0"
SAMPLE_MAX_COLUMN = 18

EVENT_COLUMNS = [
    "玩法代码",
    "玩法名称",
    "赛事ID",
    "内部ID",
    "场次编号",
    "开赛时间",
    "日期",
    "时间",
    "联赛",
    "联赛简称",
    "主队",
    "客队",
    "赛事描述",
    "赛事状态",
    "是否显示",
    "是否滚球",
    "英文联赛",
    "fixture_id",
    "主队排名",
    "客队排名",
    "来源",
]

ODDS_COLUMNS = EVENT_COLUMNS + [
    "行类型",
    "市场代码",
    "市场名称",
    "市场ID",
    "市场状态",
    "市场是否显示",
    "是否结算",
    "盘口",
    "客队盘口",
    "玩法标签",
    "选项ID",
    "选项类型",
    "选项结果",
    "选项描述",
    "选项状态",
    "选项是否显示",
    "赔率",
    "赔率分子",
    "赔率分母",
    "原始赔率",
    "变体标签",
    "变体结果",
    "变体盘口",
]


def now_ms() -> int:
    return int(time.time() * 1000)


def now() -> dt.datetime:
    return dt.datetime.now()


def format_entry_time(current_time: dt.datetime | None = None) -> str:
    return (current_time or now()).strftime("%m-%d %H:%M:%S")


# def parse_network_datetime(value: str) -> dt.datetime:
#     parsed = parsedate_to_datetime(value)
#     if parsed.tzinfo is None:
#         parsed = parsed.replace(tzinfo=dt.timezone.utc)
#     return parsed.astimezone().replace(tzinfo=None)
#
#
# def network_now(
#     *,
#     proxy_host: Any = "",
#     proxy_port: Any = "",
#     urls: Iterable[str] = NETWORK_TIME_URLS,
#     timeout: int = NETWORK_TIME_TIMEOUT_SECONDS,
# ) -> dt.datetime:
#     proxy_url = build_proxy_url(proxy_host, proxy_port)
#     handlers = []
#     if proxy_url:
#         handlers.append(urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url}))
#     opener = urllib.request.build_opener(*handlers)
#     errors: list[str] = []
#
#     for url in urls:
#         request = urllib.request.Request(
#             url,
#             headers={"User-Agent": "Mozilla/5.0"},
#             method="HEAD",
#         )
#         try:
#             with opener.open(request, timeout=timeout) as response:
#                 date_header = response.headers.get("Date")
#         except urllib.error.HTTPError as exc:
#             date_header = exc.headers.get("Date") if exc.headers else None
#             if not date_header:
#                 errors.append(f"{url}: HTTP {exc.code}")
#                 continue
#         except Exception as exc:
#             errors.append(f"{url}: {exc}")
#             continue
#
#         if not date_header:
#             errors.append(f"{url}: missing Date header")
#             continue
#
#         try:
#             return parse_network_datetime(date_header)
#         except Exception as exc:
#             errors.append(f"{url}: invalid Date header {date_header!r}: {exc}")
#
#     detail = "; ".join(errors) if errors else "no available network time source"
#     raise RuntimeError(f"network error: unable to fetch network time: {detail}")
#
#
# def resolve_run_check_time(
#     current_time: dt.datetime | None = None,
#     *,
#     proxy_host: Any = "",
#     proxy_port: Any = "",
# ) -> dt.datetime:
#     if current_time is not None:
#         return current_time
#     return network_now(proxy_host=proxy_host, proxy_port=proxy_port)


def is_run_allowed(
    current_time: dt.datetime | None = None,
    *,
    proxy_host: Any = "",
    proxy_port: Any = "",
) -> bool:
    return True


def require_run_allowed(
    current_time: dt.datetime | None = None,
    *,
    proxy_host: Any = "",
    proxy_port: Any = "",
) -> dt.datetime:
    if current_time is not None:
        return current_time
    return now()


def add_nocache(url: str) -> str:
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}nocache={now_ms()}"


def safe_text(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, (int, float, bool, dt.datetime, dt.date)):
        return value
    text = str(value)
    return text[:32767]


def decimal_odds(num: Any, den: Any, fallback: Any = None) -> Any:
    try:
        return round((int(num) + int(den)) / int(den), 6)
    except Exception:
        try:
            return float(fallback)
        except Exception:
            return ""


def clean_sheet_name(name: str, used: set[str]) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", "_", name).strip() or "Sheet"
    cleaned = cleaned[:31]
    original = cleaned
    i = 2
    while cleaned in used:
        suffix = f"_{i}"
        cleaned = f"{original[:31 - len(suffix)]}{suffix}"
        i += 1
    used.add(cleaned)
    return cleaned


def build_proxy_url(proxy_host: Any = "", proxy_port: Any = "") -> str:
    host = str(proxy_host or "").strip()
    port = str(proxy_port or "").strip()
    if not host and not port:
        return ""
    if not host or not port:
        raise ValueError("代理 IP/地址和端口号必须同时填写")
    try:
        port_number = int(port)
    except ValueError as exc:
        raise ValueError("代理端口号必须是数字") from exc
    if not 1 <= port_number <= 65535:
        raise ValueError("代理端口号必须在 1-65535 之间")
    if "://" not in host:
        host = f"http://{host}"
    return f"{host}:{port_number}"


def waf_unsbox(arg: str) -> str:
    out = [""] * len(WAF_BOX)
    for i, char in enumerate(arg):
        for j, item in enumerate(WAF_BOX):
            if item == i + 1:
                out[j] = char
                break
    return "".join(out)


def waf_hex_xor(text: str, key: str = WAF_XOR_KEY) -> str:
    out = []
    for i in range(0, min(len(text), len(key)), 2):
        out.append(f"{int(text[i:i + 2], 16) ^ int(key[i:i + 2], 16):02x}")
    return "".join(out)


def solve_waf_cookie(html: str) -> str | None:
    match = re.search(r"var\s+arg1\s*=\s*'([0-9A-Fa-f]+)'", html)
    if not match:
        return None
    return waf_hex_xor(waf_unsbox(match.group(1)))


class MacauslotClient:
    def __init__(
        self,
        entry_url: str = ENTRY_URL,
        timeout: int = 25,
        proxy_host: Any = "",
        proxy_port: Any = "",
        verify_ssl: bool = True,
    ) -> None:
        self.entry_url = entry_url
        self.timeout = timeout
        self.proxy_url = build_proxy_url(proxy_host, proxy_port)
        self.verify_ssl = verify_ssl
        self.ssl_context = None if verify_ssl else ssl._create_unverified_context()
        self.cookiejar = CookieJar()
        handlers = [urllib.request.HTTPCookieProcessor(self.cookiejar)]
        if self.ssl_context is not None:
            handlers.append(urllib.request.HTTPSHandler(context=self.ssl_context))
        if self.proxy_url:
            handlers.append(urllib.request.ProxyHandler({"http": self.proxy_url, "https": self.proxy_url}))
        self.opener = urllib.request.build_opener(*handlers)
        self.headers = {
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36"
            ),
            "Accept": "application/json,text/html,application/xhtml+xml,*/*",
            "Referer": entry_url,
        }

    def set_cookie(self, name: str, value: str, domain: str = "www.macauslot.com") -> None:
        cookie = Cookie(
            version=0,
            name=name,
            value=value,
            port=None,
            port_specified=False,
            domain=domain,
            domain_specified=True,
            domain_initial_dot=False,
            path="/",
            path_specified=True,
            secure=False,
            expires=int(time.time()) + 3600,
            discard=False,
            comment=None,
            comment_url=None,
            rest={},
            rfc2109=False,
        )
        self.cookiejar.set_cookie(cookie)

    def fetch_text(
        self,
        url: str,
        data: bytes | None = None,
        headers: dict[str, str] | None = None,
        retries: int = 3,
    ) -> str:
        merged_headers = dict(self.headers)
        if headers:
            merged_headers.update(headers)
        if data is not None:
            merged_headers.setdefault("Content-Type", "application/x-www-form-urlencoded")
            merged_headers.setdefault("Origin", BASE_URL)

        for attempt in range(retries):
            req = urllib.request.Request(url, data=data, headers=merged_headers)
            try:
                with self.opener.open(req, timeout=self.timeout) as resp:
                    raw = resp.read()
                    encoding = resp.headers.get("Content-Encoding", "").lower()
                    if "gzip" in encoding:
                        raw = gzip.decompress(raw)
                    charset = resp.headers.get_content_charset() or "utf-8"
                    text = raw.decode(charset, errors="replace")
            except urllib.error.HTTPError as exc:
                if exc.code == 404:
                    return ""
                raise

            if "acw_sc__v2" in text and "var arg1" in text:
                cookie = solve_waf_cookie(text)
                if cookie and attempt < retries - 1:
                    self.set_cookie("acw_sc__v2", cookie)
                    continue
            return text
        return text

    def get_json(self, url: str) -> Any:
        text = self.fetch_text(add_nocache(url))
        if not text.strip():
            return None
        return json.loads(text)

    def post_json(self, url: str, request_data: dict[str, str] | None = None) -> Any:
        payload = {
            "request_data": json.dumps(request_data or {}, ensure_ascii=False, separators=(",", ":")),
            "token": "",
        }
        body = urllib.parse.urlencode(payload).encode("utf-8")
        text = self.fetch_text(add_nocache(url), data=body)
        if not text.strip():
            return None
        return json.loads(text)


def realtime_feed(name: str, lang: str) -> str:
    return f"{REALTIME_URL}/{name}_{lang}_fb.json"


def extract_events(event_json: Any, lang: str, source: str, type_code: str, type_name: str) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]], list[dict[str, Any]]]:
    events: dict[str, dict[str, Any]] = {}
    configs: dict[str, dict[str, Any]] = {}
    rows: list[dict[str, Any]] = []

    if not isinstance(event_json, dict):
        return events, configs, rows

    for cfg in event_json.get("odds_config") or []:
        key = str(cfg.get("BETTING_ID", ""))
        if key:
            configs[key] = cfg

    data = event_json.get("data")
    if isinstance(data, list):
        for item in data:
            event = item.get("event") if isinstance(item, dict) else None
            if not isinstance(event, dict):
                continue
            event = dict(event)
            if item.get("event_type_en"):
                event["event_type_en"] = item.get("event_type_en")
            ev_id = str(event.get("ev_id", ""))
            if not ev_id:
                continue
            events[ev_id] = event
            rows.append(event_row(type_code, type_name, event, configs.get(ev_id), source, lang))
    elif isinstance(data, dict):
        for event in data.get("events") or []:
            if not isinstance(event, dict):
                continue
            ev_id = str(event.get("ev_id", ""))
            if not ev_id:
                continue
            events[ev_id] = event
            rows.append(event_row(type_code, type_name, event, None, source, lang))

    return events, configs, rows


def event_row(
    type_code: str,
    type_name: str,
    event: dict[str, Any],
    config: dict[str, Any] | None,
    source: str,
    lang: str,
) -> dict[str, Any]:
    config = config or {}
    event_type = event.get("eventType") or {}
    start_time = event.get("start_time", "")
    date_part, time_part = "", ""
    if isinstance(start_time, str) and " " in start_time:
        date_part, time_part = start_time.split(" ", 1)
    league_short = config.get({"cn": "TS", "sc": "SS", "en": "ES"}.get(lang, "TS"), "")
    return {
        "玩法代码": type_code,
        "玩法名称": type_name,
        "赛事ID": event.get("ev_id", ""),
        "内部ID": event.get("id", ""),
        "场次编号": event.get("shortcut", ""),
        "开赛时间": start_time,
        "日期": date_part,
        "时间": time_part[:5],
        "联赛": event_type.get("name", "") or event.get("event_type", ""),
        "联赛简称": league_short,
        "主队": event.get("home_team", ""),
        "客队": event.get("away_team", ""),
        "赛事描述": event.get("desc", ""),
        "赛事状态": event.get("state", event.get("status", "")),
        "是否显示": event.get("displayed", ""),
        "是否滚球": event.get("has_bet_in_run", event.get("bet_in_run", "")),
        "英文联赛": event.get("event_type_en", ""),
        "fixture_id": config.get("FIXTURE_ID", ""),
        "主队排名": config.get("HR", ""),
        "客队排名": config.get("AR", ""),
        "来源": source,
    }


def flatten_market_rows(
    type_code: str,
    type_name: str,
    odds_json: Any,
    events: dict[str, dict[str, Any]],
    configs: dict[str, dict[str, Any]],
    source: str,
    lang: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if not isinstance(odds_json, dict):
        return rows
    for item in odds_json.get("data") or []:
        if not isinstance(item, dict):
            continue
        ev_id = str(item.get("ev_id", ""))
        event = events.get(ev_id, {"ev_id": ev_id})
        config = configs.get(ev_id)
        rows.extend(flatten_markets_for_event(type_code, type_name, event, config, item, source, lang))
    return rows


def flatten_embedded_event_rows(
    type_code: str,
    type_name: str,
    events: Iterable[dict[str, Any]],
    source: str,
    lang: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    event_rows: list[dict[str, Any]] = []
    odds_rows: list[dict[str, Any]] = []
    for event in events:
        if not isinstance(event, dict):
            continue
        event_rows.append(event_row(type_code, type_name, event, None, source, lang))
        odds_rows.extend(flatten_markets_for_event(type_code, type_name, event, None, event, source, lang))
    return event_rows, odds_rows


def flatten_markets_for_event(
    type_code: str,
    type_name: str,
    event: dict[str, Any],
    config: dict[str, Any] | None,
    market_holder: dict[str, Any],
    source: str,
    lang: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    base = event_row(type_code, type_name, event, config, source, lang)
    for market in market_holder.get("markets") or []:
        if not isinstance(market, dict):
            continue
        market_base = {
            **base,
            "市场代码": market.get("sort", ""),
            "市场名称": (market.get("outcomeGroup") or {}).get("name", "") or market.get("name", ""),
            "市场ID": market.get("ev_mkt_id", ""),
            "市场状态": market.get("status", ""),
            "市场是否显示": market.get("displayed", ""),
            "是否结算": market.get("settled", ""),
            "盘口": market.get("hcap_disp", ""),
            "客队盘口": market.get("away_hcap_disp", ""),
            "玩法标签": (market.get("outcomeGroup") or {}).get("tag", ""),
        }
        added = False
        for outcome in market.get("outcomes") or []:
            if not isinstance(outcome, dict):
                continue
            lp_num = outcome.get("lp_num")
            lp_den = outcome.get("lp_den")
            rows.append(
                {
                    **market_base,
                    "行类型": "outcome",
                    "选项ID": outcome.get("ev_oc_id", ""),
                    "选项类型": outcome.get("type", ""),
                    "选项结果": outcome.get("fb_result", ""),
                    "选项描述": outcome.get("desc", ""),
                    "选项状态": outcome.get("status", ""),
                    "选项是否显示": outcome.get("displayed", ""),
                    "赔率": decimal_odds(lp_num, lp_den, outcome.get("rate", "")),
                    "赔率分子": lp_num,
                    "赔率分母": lp_den,
                    "原始赔率": outcome.get("rate", ""),
                }
            )
            added = True
        for variant in market.get("marketVariants") or []:
            if not isinstance(variant, dict):
                continue
            ov = variant.get("outcomeVariant") or {}
            lp_num = ov.get("price_num")
            lp_den = ov.get("price_den")
            rows.append(
                {
                    **market_base,
                    "行类型": "marketVariant",
                    "选项ID": ov.get("ev_oc_id", ""),
                    "选项类型": variant.get("type", ""),
                    "选项结果": variant.get("fb_result", ""),
                    "选项描述": variant.get("desc", ""),
                    "选项状态": variant.get("status", ""),
                    "选项是否显示": variant.get("displayed", ""),
                    "赔率": decimal_odds(lp_num, lp_den, ov.get("rate", "")),
                    "赔率分子": lp_num,
                    "赔率分母": lp_den,
                    "原始赔率": ov.get("rate", ""),
                    "变体标签": variant.get("tag", ""),
                    "变体结果": variant.get("fb_result", ""),
                    "变体盘口": variant.get("hcap_disp", ""),
                }
            )
            added = True
        if not added:
            rows.append({**market_base, "行类型": "market"})
    return rows


def event_ids(events: dict[str, dict[str, Any]]) -> str:
    ids = [str(ev_id) for ev_id in events.keys() if ev_id]
    return ",".join(ids) + ("," if ids else "")


@dataclass
class ScrapeResult:
    type_code: str
    type_name: str
    source: str
    event_rows: list[dict[str, Any]]
    odds_rows: list[dict[str, Any]]
    status: str = "ok"
    note: str = ""


@dataclass
class RunScrapeResult:
    output: Path
    updated_count: int = 0
    appended_count: int = 0

    @property
    def changed(self) -> bool:
        return self.updated_count > 0 or self.appended_count > 0

    def __fspath__(self) -> str:
        return str(self.output)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Path):
            return self.output == other
        if isinstance(other, RunScrapeResult):
            return (
                self.output == other.output
                and self.updated_count == other.updated_count
                and self.appended_count == other.appended_count
            )
        return False


def scrape_type(client: MacauslotClient, type_code: str, lang: str) -> ScrapeResult:
    type_name = ODDS_TYPE_NAMES.get(type_code, type_code)

    if type_code == "index":
        prefix = "threeinone"
        event_url = realtime_feed(f"{prefix}_event", lang)
        odds_url = realtime_feed(f"{prefix}_odds", lang)
        event_json = client.get_json(event_url)
        events, configs, event_rows = extract_events(event_json, lang, event_url, type_code, type_name)
        odds_json = client.post_json(odds_url, {"event": event_ids(events)}) if events else None
        odds_rows = flatten_market_rows(type_code, type_name, odds_json, events, configs, odds_url, lang)
        return ScrapeResult(type_code, type_name, f"{event_url} | {odds_url}", event_rows, odds_rows)

    if type_code == "half_index":
        prefix = "firsthalf"
        event_url = realtime_feed(f"{prefix}_event", lang)
        odds_url = realtime_feed(f"{prefix}_odds", lang)
        event_json = client.get_json(event_url)
        events, configs, event_rows = extract_events(event_json, lang, event_url, type_code, type_name)
        odds_json = client.post_json(odds_url, {"event": event_ids(events)}) if events else None
        odds_rows = flatten_market_rows(type_code, type_name, odds_json, events, configs, odds_url, lang)
        return ScrapeResult(type_code, type_name, f"{event_url} | {odds_url}", event_rows, odds_rows)

    if type_code == "bir":
        source = realtime_feed("realtime_all", lang)
        data = client.get_json(source)
        event_rows: list[dict[str, Any]] = []
        odds_rows: list[dict[str, Any]] = []
        if isinstance(data, list) and data:
            bundle = data[0]
            events = bundle.get("events") or []
            event_rows, _ = flatten_embedded_event_rows(type_code, type_name, events, source, lang)
            event_map = {str(event.get("ev_id")): event for event in events if isinstance(event, dict)}
            markets = bundle.get("markets") or {}
            odds_rows = flatten_market_rows(type_code, type_name, markets, event_map, {}, source, lang)
        return ScrapeResult(type_code, type_name, source, event_rows, odds_rows)

    if type_code.startswith("FUNC_"):
        code = type_code.replace("FUNC_", "")
        source = realtime_feed(code, lang)
        data = client.get_json(source)
        events = []
        if isinstance(data, dict) and isinstance(data.get("data"), dict):
            events = data["data"].get("events") or []
        event_rows, odds_rows = flatten_embedded_event_rows(type_code, type_name, events, source, lang)
        return ScrapeResult(type_code, type_name, source, event_rows, odds_rows)

    feed_code = JSON_CODE_MAP[type_code]
    event_url = realtime_feed(f"{feed_code}_event", lang)
    odds_url = realtime_feed(f"{feed_code}_odds", lang)
    event_json = client.get_json(event_url)
    events, configs, event_rows = extract_events(event_json, lang, event_url, type_code, type_name)
    request_key = "events"
    odds_json = client.post_json(odds_url, {request_key: event_ids(events)}) if events else None
    odds_rows = flatten_market_rows(type_code, type_name, odds_json, events, configs, odds_url, lang)
    note = "" if events else "该接口当前无赛事数据"
    return ScrapeResult(type_code, type_name, f"{event_url} | {odds_url}", event_rows, odds_rows, note=note)


def discover_special_types(client: MacauslotClient) -> list[str]:
    url = f"{REALTIME_URL}/menu_fb.json"
    try:
        menu = client.get_json(url)
    except Exception:
        return []
    if not isinstance(menu, list):
        return []
    return [f"FUNC_{item}" for item in menu if item in SPECIAL_MENU_IDS]


def write_sheet(ws, rows: list[dict[str, Any]], columns: list[str]) -> None:
    for col_idx, header in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(rows, 2):
        for col_idx, header in enumerate(columns, 1):
            ws.cell(row=row_idx, column=col_idx, value=safe_text(row.get(header, "")))

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col_idx, header in enumerate(columns, 1):
        max_len = len(str(header))
        for row_idx in range(2, min(ws.max_row, 80) + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                max_len = max(max_len, min(len(str(value)), 60))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 32)


def build_workbook(results: list[ScrapeResult], output: Path, lang: str, entry_url: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "抓取概览"

    created_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    overview_rows = [
        {"字段": "来源页面", "内容": entry_url},
        {"字段": "抓取时间", "内容": created_at},
        {"字段": "语言参数", "内容": lang},
        {"字段": "说明", "内容": "每个赔率工作表为页面下拉项对应 JSON 数据的扁平化表格；一行代表一个赔率选项或盘口变体。"},
    ]
    write_sheet(ws, overview_rows, ["字段", "内容"])

    used_sheet_names = {ws.title}
    summary_rows: list[dict[str, Any]] = []
    all_event_rows: list[dict[str, Any]] = []

    for result in results:
        all_event_rows.extend(result.event_rows)
        sheet_name = clean_sheet_name(result.type_name, used_sheet_names)
        summary_rows.append(
            {
                "工作表": sheet_name,
                "玩法代码": result.type_code,
                "玩法名称": result.type_name,
                "事件数": len(result.event_rows),
                "赔率行数": len(result.odds_rows),
                "状态": result.status,
                "数据源": result.source,
                "备注": result.note,
            }
        )
        ws_data = wb.create_sheet(sheet_name)
        if result.odds_rows:
            extra = sorted({key for row in result.odds_rows for key in row.keys()} - set(ODDS_COLUMNS))
            write_sheet(ws_data, result.odds_rows, ODDS_COLUMNS + extra)
        else:
            write_sheet(ws_data, [{"提示": result.note or "当前无数据"}], ["提示"])

    ws_summary = wb.create_sheet("接口汇总", 1)
    write_sheet(ws_summary, summary_rows, SUMMARY_COLUMNS)

    ws_events = wb.create_sheet("全部赛事", 2)
    if all_event_rows:
        seen = set()
        unique_events = []
        for row in all_event_rows:
            key = (row.get("玩法代码"), row.get("赛事ID"))
            if key not in seen:
                seen.add(key)
                unique_events.append(row)
        write_sheet(ws_events, unique_events, EVENT_COLUMNS)
    else:
        write_sheet(ws_events, [{"提示": "当前无赛事数据"}], ["提示"])

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def visible_handicap(row: dict[str, Any]) -> Any:
    if row.get("变体盘口"):
        return row.get("变体盘口")
    if row.get("选项类型") == "A" and row.get("客队盘口"):
        return row.get("客队盘口")
    return row.get("盘口", "")


def visible_selection(row: dict[str, Any]) -> str:
    desc = str(row.get("选项描述") or "").strip()
    if desc:
        return desc

    market_code = str(row.get("市场代码") or "").upper()
    result = str(row.get("变体结果") or row.get("选项结果") or "").upper()
    home = str(row.get("主队") or "")
    away = str(row.get("客队") or "")

    if market_code in {"AH", "A1", "CH", "EA", "PA"}:
        if result in {"H", "HH", "HL"}:
            return home
        if result in {"A", "AH", "AL"}:
            return away
    if market_code in {"HL", "H1", "CO", "C1", "EH", "PH", "EV"}:
        if result in {"H", "HH", "HL"}:
            return "上"
        if result in {"A", "AH", "AL", "L", "LH", "LL"}:
            return "下"
    if result == "D":
        return "和"
    return result or str(row.get("选项类型") or "")


def to_visible_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    visible: list[dict[str, Any]] = []
    for row in rows:
        if not row.get("赔率"):
            continue
        start_time = str(row.get("开赛时间") or "")
        display_time = str(row.get("时间") or "")
        if not display_time and " " in start_time:
            display_time = start_time.split(" ", 1)[1][:5]
        league = row.get("联赛简称") or row.get("联赛") or ""
        market = row.get("市场名称") or row.get("市场代码") or ""
        visible.append(
            {
                "日期": row.get("日期", ""),
                "类别": league,
                "开赛时间": display_time,
                "主/客": f"{row.get('主队', '')}\n{row.get('客队', '')}".strip(),
                "市场": market,
                "选项": visible_selection(row),
                "盘口": visible_handicap(row),
                "赔率": row.get("赔率", ""),
            }
        )
    return visible


def build_visible_workbook(results: list[ScrapeResult], output: Path, lang: str, entry_url: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "抓取概览"

    created_at = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    overview_rows = [
        {"字段": "来源页面", "内容": entry_url},
        {"字段": "抓取时间", "内容": created_at},
        {"字段": "语言参数", "内容": lang},
        {"字段": "说明", "内容": "网页样式版：日期为合并分隔行，同一场比赛的类别、开赛时间、主/客以合并单元格显示。"},
    ]
    write_sheet(ws, overview_rows, ["字段", "内容"])

    used_sheet_names = {ws.title}
    summary_rows: list[dict[str, Any]] = []

    for result in results:
        visible_rows = to_visible_rows(result.odds_rows)
        sheet_name = clean_sheet_name(result.type_name, used_sheet_names)
        summary_rows.append(
            {
                "工作表": sheet_name,
                "玩法名称": result.type_name,
                "事件数": len(result.event_rows),
                "赔率行数": len(visible_rows),
                "状态": result.status,
                "备注": result.note,
            }
        )
        ws_data = wb.create_sheet(sheet_name)
        if visible_rows:
            write_web_style_sheet(ws_data, visible_rows)
        else:
            write_sheet(ws_data, [{"提示": result.note or "当前网页无此玩法数据"}], ["提示"])

    ws_summary = wb.create_sheet("接口汇总", 1)
    write_sheet(ws_summary, summary_rows, SUMMARY_COLUMNS)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def write_web_style_sheet(ws, rows: list[dict[str, Any]]) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    date_fill = PatternFill("solid", fgColor="1F644F")
    event_fill = PatternFill("solid", fgColor="EAF4EF")
    market_fill = PatternFill("solid", fgColor="F3F6F8")
    thin = Side(style="thin", color="D9E2E8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, header in enumerate(VISIBLE_COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    widths = [16, 12, 28, 18, 34, 12, 10]
    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    current_row = 2
    current_date: Any = object()
    i = 0
    while i < len(rows):
        row = rows[i]
        row_date = row.get("日期", "")
        if row_date != current_date:
            current_date = row_date
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(VISIBLE_COLUMNS))
            cell = ws.cell(row=current_row, column=1, value=safe_text(row_date))
            cell.fill = date_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for col_idx in range(1, len(VISIBLE_COLUMNS) + 1):
                ws.cell(row=current_row, column=col_idx).border = border
                ws.cell(row=current_row, column=col_idx).fill = date_fill
            current_row += 1

        event_key = (row.get("日期", ""), row.get("类别", ""), row.get("开赛时间", ""), row.get("主/客", ""))
        group: list[dict[str, Any]] = []
        while i < len(rows):
            candidate = rows[i]
            candidate_key = (
                candidate.get("日期", ""),
                candidate.get("类别", ""),
                candidate.get("开赛时间", ""),
                candidate.get("主/客", ""),
            )
            if candidate_key != event_key:
                break
            group.append(candidate)
            i += 1

        group_start = current_row
        for item in group:
            values = [
                item.get("类别", ""),
                item.get("开赛时间", ""),
                item.get("主/客", ""),
                item.get("市场", ""),
                item.get("选项", ""),
                item.get("盘口", ""),
                item.get("赔率", ""),
            ]
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=safe_text(value))
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if col_idx <= 3:
                    cell.fill = event_fill
                elif col_idx == 4:
                    cell.fill = market_fill
            current_row += 1

        group_end = current_row - 1
        if group_end > group_start:
            for col_idx in (1, 2, 3):
                ws.merge_cells(start_row=group_start, start_column=col_idx, end_row=group_end, end_column=col_idx)
                ws.cell(row=group_start, column=col_idx).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

        market_start = group_start
        while market_start <= group_end:
            market_value = ws.cell(row=market_start, column=4).value
            market_end = market_start
            while market_end + 1 <= group_end and ws.cell(row=market_end + 1, column=4).value == market_value:
                market_end += 1
            if market_end > market_start:
                ws.merge_cells(start_row=market_start, start_column=4, end_row=market_end, end_column=4)
                ws.cell(row=market_start, column=4).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            market_start = market_end + 1

    for row_idx in range(1, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 24
    for row_idx in range(2, ws.max_row + 1):
        if ws.cell(row=row_idx, column=1).value and ws.cell(row=row_idx, column=1).fill.fgColor.rgb == "001F644F":
            ws.row_dimensions[row_idx].height = 22


def visible_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return (
        row.get("日期", ""),
        row.get("类别", ""),
        row.get("开赛时间", ""),
        row.get("主/客", ""),
        row.get("市场", ""),
        row.get("选项", ""),
        row.get("盘口", ""),
    )


def is_date_divider(values: list[Any]) -> bool:
    return bool(values and values[0] and all(value in (None, "") for value in values[1:]))


def extract_web_rows_with_positions(ws) -> dict[tuple[Any, ...], int]:
    positions: dict[tuple[Any, ...], int] = {}
    current_date = ""
    current_league = ""
    current_time = ""
    current_match = ""
    current_market = ""

    for row_idx in range(2, ws.max_row + 1):
        values = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, 8)]
        if is_date_divider(values):
            current_date = values[0]
            continue

        if values[0] not in (None, ""):
            current_league = values[0]
        if values[1] not in (None, ""):
            current_time = values[1]
        if values[2] not in (None, ""):
            current_match = values[2]
        if values[3] not in (None, ""):
            current_market = values[3]

        row = {
            "日期": current_date,
            "类别": current_league,
            "开赛时间": current_time,
            "主/客": current_match,
            "市场": current_market,
            "选项": values[4] or "",
            "盘口": values[5] or "",
            "赔率": values[6] or "",
        }
        if row["选项"] or row["赔率"]:
            positions.setdefault(visible_key(row), row_idx)

    return positions


def append_web_style_rows(ws, rows: list[dict[str, Any]]) -> None:
    if not rows:
        return

    date_fill = PatternFill("solid", fgColor="1F644F")
    event_fill = PatternFill("solid", fgColor="EAF4EF")
    market_fill = PatternFill("solid", fgColor="F3F6F8")
    thin = Side(style="thin", color="D9E2E8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    current_row = ws.max_row + 1
    current_date: Any = object()
    i = 0
    while i < len(rows):
        row = rows[i]
        row_date = row.get("日期", "")
        if row_date != current_date:
            current_date = row_date
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(VISIBLE_COLUMNS))
            cell = ws.cell(row=current_row, column=1, value=safe_text(row_date))
            cell.fill = date_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            for col_idx in range(1, len(VISIBLE_COLUMNS) + 1):
                ws.cell(row=current_row, column=col_idx).border = border
                ws.cell(row=current_row, column=col_idx).fill = date_fill
            ws.row_dimensions[current_row].height = 22
            current_row += 1

        event_key = (row.get("日期", ""), row.get("类别", ""), row.get("开赛时间", ""), row.get("主/客", ""))
        group: list[dict[str, Any]] = []
        while i < len(rows):
            candidate = rows[i]
            candidate_key = (
                candidate.get("日期", ""),
                candidate.get("类别", ""),
                candidate.get("开赛时间", ""),
                candidate.get("主/客", ""),
            )
            if candidate_key != event_key:
                break
            group.append(candidate)
            i += 1

        group_start = current_row
        for item in group:
            values = [
                item.get("类别", ""),
                item.get("开赛时间", ""),
                item.get("主/客", ""),
                item.get("市场", ""),
                item.get("选项", ""),
                item.get("盘口", ""),
                item.get("赔率", ""),
            ]
            for col_idx, value in enumerate(values, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=safe_text(value))
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                if col_idx <= 3:
                    cell.fill = event_fill
                elif col_idx == 4:
                    cell.fill = market_fill
            ws.row_dimensions[current_row].height = 24
            current_row += 1

        group_end = current_row - 1
        if group_end > group_start:
            for col_idx in (1, 2, 3):
                ws.merge_cells(start_row=group_start, start_column=col_idx, end_row=group_end, end_column=col_idx)
                ws.cell(row=group_start, column=col_idx).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

        market_start = group_start
        while market_start <= group_end:
            market_value = ws.cell(row=market_start, column=4).value
            market_end = market_start
            while market_end + 1 <= group_end and ws.cell(row=market_end + 1, column=4).value == market_value:
                market_end += 1
            if market_end > market_start:
                ws.merge_cells(start_row=market_start, start_column=4, end_row=market_end, end_column=4)
                ws.cell(row=market_start, column=4).alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )
            market_start = market_end + 1


def update_single_sheet_workbook(output: Path, visible_rows: list[dict[str, Any]]) -> tuple[int, int]:
    sheet_name = ODDS_TYPE_NAMES["index"]

    if output.exists():
        wb = load_workbook(output)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active
            ws.title = sheet_name
            write_web_style_sheet(ws, [])
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        write_web_style_sheet(ws, visible_rows)
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
        return 0, len(visible_rows)

    existing_positions = extract_web_rows_with_positions(ws)
    appended_rows: list[dict[str, Any]] = []
    updated_count = 0

    for row in visible_rows:
        key = visible_key(row)
        row_idx = existing_positions.get(key)
        if row_idx:
            current_value = ws.cell(row=row_idx, column=7).value
            new_value = row.get("赔率", "")
            if current_value != new_value:
                ws.cell(row=row_idx, column=7, value=safe_text(new_value))
                updated_count += 1
        else:
            appended_rows.append(row)
            existing_positions[key] = -1

    append_web_style_rows(ws, appended_rows)

    # Keep the workbook to one worksheet in the maintained file.
    for name in list(wb.sheetnames):
        if name != sheet_name:
            del wb[name]

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return updated_count, len(appended_rows)


def format_odds_value(value: Any) -> str:
    if value in (None, ""):
        return ""
    try:
        number = float(value)
        text = f"{number:.6f}".rstrip("0").rstrip(".")
        return text
    except Exception:
        return str(value)


def sample_value(label: str, hcap: Any, odds: Any) -> str:
    odds_text = format_odds_value(odds)
    if not odds_text:
        return "-"
    hcap_text = "" if hcap in (None, "") else str(hcap)
    return f"{label}{hcap_text}{odds_text}"


def sample_market_value(label: str, hcap: Any, odds: Any) -> dict[str, Any]:
    return {"label": label, "盘口": hcap if hcap not in (None, "") else "-", "赔率": odds}


def split_sample_market_value(value: Any) -> tuple[Any, Any]:
    if isinstance(value, dict):
        label = str(value.get("label") or "")
        hcap = value.get("盘口", "")
        odds = value.get("赔率", "")
        hcap_text = "-" if hcap in (None, "") else str(hcap)
        display_hcap = hcap_text if not label or hcap_text == "-" else f"{label}{hcap_text}"
        return display_hcap, odds if odds not in (None, "") else None
    return value, None


def sample_date_title(date_text: Any) -> str:
    text = str(date_text or "")
    try:
        date_obj = dt.datetime.strptime(text, "%Y-%m-%d").date()
    except Exception:
        return text
    weekdays = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    return f"{date_obj.year} 年{date_obj.month}月 {date_obj.day} 日({weekdays[date_obj.weekday()]})"


def sample_date_value_from_title(value: Any) -> Any:
    match = re.search(r"(\d{4})\s*年\s*(\d{1,2})月\s*(\d{1,2})", str(value or ""))
    if not match:
        return value
    year, month, day = match.groups()
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def sample_key(match: dict[str, Any]) -> tuple[Any, ...]:
    return (
        match.get("日期", ""),
        match.get("类别", ""),
        match.get("比赛时间", ""),
        match.get("主/客", ""),
    )


def raw_to_sample_matches(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    matches: dict[tuple[Any, ...], dict[str, Any]] = {}
    order: list[tuple[Any, ...]] = []

    def get_match(row: dict[str, Any]) -> dict[str, Any]:
        start_time = str(row.get("开赛时间") or "")
        display_time = str(row.get("时间") or "")
        if not display_time and " " in start_time:
            display_time = start_time.split(" ", 1)[1][:5]
        match = {
            "日期": row.get("日期", ""),
            "类别": row.get("联赛简称") or row.get("联赛") or "",
            "比赛时间": display_time,
            "主/客": f"{row.get('主队', '')}\n{row.get('客队', '')}".strip(),
            "让球主": [sample_market_value("", "-", "") for _ in range(3)],
            "让球客": [sample_market_value("", "-", "") for _ in range(3)],
            "上下上": [sample_market_value("", "-", "") for _ in range(3)],
            "上下下": [sample_market_value("", "-", "") for _ in range(3)],
            "标准盘": ["-", "-", "-"],
            "比分": ["", "", ""],
        }
        key = sample_key(match)
        if key not in matches:
            matches[key] = match
            order.append(key)
        return matches[key]

    for row in rows:
        if not row.get("赔率"):
            continue
        market = str(row.get("市场代码") or "").upper()
        match = get_match(row)
        row_type = str(row.get("行类型") or "")
        result = str(row.get("变体结果") or row.get("选项结果") or "").upper()
        option_type = str(row.get("选项类型") or "").upper()
        tag = str(row.get("变体标签") or "").upper()
        odds = row.get("赔率", "")

        if market == "MR":
            idx = {"H": 0, "D": 1, "A": 2}.get(result)
            label = {"H": "主", "D": "和", "A": "客"}.get(result, "")
            if idx is not None:
                match["标准盘"][idx] = sample_value(label, "", odds)

        elif market == "AH":
            idx = 1
            if row_type == "marketVariant":
                idx = 0 if tag == "HI1" else 2 if tag == "LO1" else 1
            side = "home" if (option_type == "H" or result in {"H", "HH", "HL"}) else "away"
            hcap = visible_handicap(row)
            if side == "home":
                match["让球主"][idx] = sample_market_value("", hcap, odds)
            else:
                match["让球客"][idx] = sample_market_value("", hcap, odds)

        elif market == "HL":
            idx = 1
            if row_type == "marketVariant":
                idx = 0 if tag == "HI1" else 2 if tag == "LO1" else 1
            is_over = option_type == "H" or result in {"H", "HH", "HL"}
            hcap = visible_handicap(row)
            if is_over:
                match["上下上"][idx] = sample_market_value("上", hcap, odds)
            else:
                match["上下下"][idx] = sample_market_value("下", hcap, odds)

    return [matches[key] for key in order]


def write_sample_header(ws, row_idx: int) -> None:
    for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
        value = SAMPLE_HEADERS.get(col_idx, "")
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        apply_header_style(cell, col_idx)
    ws.merge_cells(start_row=row_idx, start_column=4, end_row=row_idx, end_column=9)
    ws.merge_cells(start_row=row_idx, start_column=11, end_row=row_idx, end_column=14)


def sample_style_for(row: int, col: int) -> dict[str, Any] | None:
    if not SAMPLE_TEMPLATE_PATH.exists():
        return None
    if not hasattr(sample_style_for, "_cache"):
        wb = load_workbook(SAMPLE_TEMPLATE_PATH)
        ws = wb.active
        cache: dict[tuple[int, int], dict[str, Any]] = {}
        for r in range(1, min(ws.max_row, 5) + 1):
            for c in range(1, min(ws.max_column, SAMPLE_MAX_COLUMN) + 1):
                src = ws.cell(r, c)
                cache[(r, c)] = {
                    "font": copy.copy(src.font),
                    "fill": copy.copy(src.fill),
                    "border": copy.copy(src.border),
                    "alignment": copy.copy(src.alignment),
                    "number_format": src.number_format,
                    "protection": copy.copy(src.protection),
                }
        sample_style_for._cache = cache  # type: ignore[attr-defined]
    return sample_style_for._cache.get((row, col))  # type: ignore[attr-defined]


def apply_template_style(cell, row: int, col: int) -> bool:
    if isinstance(cell, MergedCell):
        return True
    style = sample_style_for(row, col)
    if not style:
        return False
    cell.font = copy.copy(style["font"])
    cell.fill = copy.copy(style["fill"])
    cell.border = copy.copy(style["border"])
    cell.alignment = copy.copy(style["alignment"])
    cell.number_format = style["number_format"]
    cell.protection = copy.copy(style["protection"])
    return True


def apply_header_style(cell, col_idx: int) -> None:
    if apply_template_style(cell, 2, col_idx):
        return
    thin = Side(style="thin", color="000000")
    cell.fill = PatternFill("solid", fgColor="FFD966")
    cell.font = Font(name="宋体", size=11, bold=True, color="08090C")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_body_style(cell, offset: int, col_idx: int) -> None:
    if apply_template_style(cell, 3 + offset, col_idx):
        return
    thin = Side(style="thin", color="000000")
    cell.fill = PatternFill("solid", fgColor="FFFFFF")
    cell.font = Font(name="宋体", size=11, bold=False, color="000000")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_date_style(cell, col_idx: int = 1) -> None:
    if apply_template_style(cell, 1, col_idx):
        return
    thin = Side(style="thin", color="000000")
    cell.font = Font(name="宋体", size=11, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def setup_sample_sheet(ws) -> None:
    ws.sheet_view.showGridLines = False
    widths = [20.375] + [13] * (SAMPLE_MAX_COLUMN - 1)
    if SAMPLE_TEMPLATE_PATH.exists():
        tpl = load_workbook(SAMPLE_TEMPLATE_PATH).active
        widths = [
            tpl.column_dimensions[get_column_letter(col_idx)].width or widths[col_idx - 1]
            for col_idx in range(1, SAMPLE_MAX_COLUMN + 1)
        ]
    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = "A3"


def append_sample_matches(
    ws,
    matches: list[dict[str, Any]],
    include_header: bool = True,
    entry_time: str | None = None,
) -> None:
    setup_sample_sheet(ws)
    has_content = ws.max_row > 1 or ws.cell(1, 1).value not in (None, "")
    current_row = ws.max_row + 1 if has_content else 1
    current_date: Any = None

    for match in matches:
        date_text = match.get("日期", "")
        if date_text != current_date:
            current_date = date_text
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=SAMPLE_MAX_COLUMN)
            ws.cell(row=current_row, column=1, value=sample_date_title(date_text))
            for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
                apply_date_style(ws.cell(row=current_row, column=col_idx), col_idx)
            ws.row_dimensions[current_row].height = 24
            current_row += 1
            if include_header:
                write_sample_header(ws, current_row)
                ws.row_dimensions[current_row].height = 25
                current_row += 1

        start = current_row
        for offset in range(3):
            row_idx = start + offset
            home_hcap, home_odds = split_sample_market_value(match["让球主"][offset])
            away_hcap, away_odds = split_sample_market_value(match["让球客"][offset])
            over_hcap, over_odds = split_sample_market_value(match["上下上"][offset])
            under_hcap, under_odds = split_sample_market_value(match["上下下"][offset])
            values = {
                1: match.get("类别", "") if offset == 0 else "",
                2: match.get("比赛时间", "") if offset == 0 else "",
                3: match.get("主/客", "") if offset == 0 else "",
                4: "主" if offset == 0 else "",
                5: home_hcap,
                6: home_odds,
                7: "客" if offset == 1 else "",
                8: away_hcap,
                9: away_odds,
                10: entry_time or "",
                11: over_hcap,
                12: over_odds,
                13: under_hcap,
                14: under_odds,
                15: entry_time or "",
                16: match["标准盘"][offset],
                17: entry_time or "",
                18: match["比分"][offset],
            }
            for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=values.get(col_idx, ""))
                apply_body_style(cell, offset, col_idx)
            ws.row_dimensions[row_idx].height = [25, 24, 29][offset]

        for col_idx in (1, 2, 3, 4):
            ws.merge_cells(start_row=start, start_column=col_idx, end_row=start + 2, end_column=col_idx)
            ws.cell(row=start, column=col_idx).alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
        current_row += 3


def create_sample_workbook(output: Path, matches: list[dict[str, Any]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SAMPLE_SHEET_NAME
    entry_time = format_entry_time()
    append_sample_matches(ws, matches, include_header=True, entry_time=entry_time)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def normalize_sample_styles(ws) -> None:
    setup_sample_sheet(ws)
    row_idx = 1
    while row_idx <= ws.max_row:
        values = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, SAMPLE_MAX_COLUMN + 1)]
        if values[0] and all(value in (None, "") for value in values[1:]):
            for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
                apply_date_style(ws.cell(row=row_idx, column=col_idx), col_idx)
            ws.row_dimensions[row_idx].height = 24
            row_idx += 1
            continue
        if values[0] == "类别":
            for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
                apply_header_style(ws.cell(row=row_idx, column=col_idx), col_idx)
            ws.row_dimensions[row_idx].height = 25
            row_idx += 1
            continue
        if values[0] and values[1] and values[2]:
            for offset in range(3):
                target_row = row_idx + offset
                if target_row > ws.max_row:
                    break
                for col_idx in range(1, SAMPLE_MAX_COLUMN + 1):
                    apply_body_style(ws.cell(row=target_row, column=col_idx), offset, col_idx)
                ws.row_dimensions[target_row].height = [25, 24, 29][offset]
            row_idx += 3
            continue
        row_idx += 1


def is_sample_sheet(ws) -> bool:
    return (
        ws.max_column >= SAMPLE_MAX_COLUMN
        and ws.cell(2, 1).value == "类别"
        and ws.cell(2, 4).value == "让球盘"
        and ws.cell(2, 11).value == "上/下盘"
    )


def extract_sample_positions(ws) -> dict[tuple[Any, ...], int]:
    positions: dict[tuple[Any, ...], int] = {}
    current_date = ""
    row_idx = 1
    while row_idx <= ws.max_row:
        values = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, SAMPLE_MAX_COLUMN + 1)]
        if values[0] and all(value in (None, "") for value in values[1:]):
            current_date = sample_date_value_from_title(values[0])
            row_idx += 1
            continue
        if values[0] == "类别":
            row_idx += 1
            continue
        if values[0] and values[1] and values[2]:
            key = (current_date, values[0], values[1], values[2])
            positions[key] = row_idx
            row_idx += 3
        else:
            row_idx += 1
    return positions


def update_sample_workbook(output: Path, matches: list[dict[str, Any]]) -> tuple[int, int]:
    if not output.exists():
        create_sample_workbook(output, matches)
        return 0, len(matches)

    wb = load_workbook(output)
    ws = wb[SAMPLE_SHEET_NAME] if SAMPLE_SHEET_NAME in wb.sheetnames else wb.active
    ws.title = SAMPLE_SHEET_NAME

    if not is_sample_sheet(ws):
        # One-time conversion from an older generated layout to the requested sample layout.
        wb = Workbook()
        ws = wb.active
        ws.title = SAMPLE_SHEET_NAME
        append_sample_matches(ws, matches, include_header=True, entry_time=format_entry_time())
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
        return 0, len(matches)

    positions = extract_sample_positions(ws)
    appended: list[dict[str, Any]] = []
    updated = 0
    update_time = format_entry_time()

    for match in matches:
        key = sample_key(match)
        start = positions.get(key)
        if start:
            update_map = []
            for offset in range(3):
                home_hcap, home_odds = split_sample_market_value(match["让球主"][offset])
                away_hcap, away_odds = split_sample_market_value(match["让球客"][offset])
                over_hcap, over_odds = split_sample_market_value(match["上下上"][offset])
                under_hcap, under_odds = split_sample_market_value(match["上下下"][offset])
                update_map.extend(
                    [
                        (offset, 5, home_hcap, 10),
                        (offset, 6, home_odds, 10),
                        (offset, 8, away_hcap, 10),
                        (offset, 9, away_odds, 10),
                        (offset, 11, over_hcap, 15),
                        (offset, 12, over_odds, 15),
                        (offset, 13, under_hcap, 15),
                        (offset, 14, under_odds, 15),
                        (offset, 16, match["标准盘"][offset], 17),
                    ]
                )
            for offset, col_idx, value, timestamp_col_idx in update_map:
                cell = ws.cell(row=start + offset, column=col_idx)
                if cell.value != value:
                    cell.value = value
                    ws.cell(row=start + offset, column=timestamp_col_idx, value=update_time)
                    updated += 1
        else:
            appended.append(match)
            positions[key] = -1

    append_sample_matches(ws, appended, include_header=True, entry_time=update_time)
    normalize_sample_styles(ws)

    for name in list(wb.sheetnames):
        if name != SAMPLE_SHEET_NAME:
            del wb[name]

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return updated, len(appended)


def require_visible_scrape_data(results: list[ScrapeResult]) -> None:
    if not results:
        raise RuntimeError("没有抓取到任何玩法数据，已取消写入 Excel")
    result = results[0]
    if result.status != "ok":
        detail = f"：{result.note}" if result.note else ""
        raise RuntimeError(f"{result.type_name} 抓取失败{detail}，已取消写入 Excel")
    if not result.odds_rows:
        raise RuntimeError(f"{result.type_name} 未抓取到赔率数据，已取消写入 Excel")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape Macauslot soccer odds and save to Excel.")
    parser.add_argument("--url", default=ENTRY_URL, help="Macauslot odds entry URL.")
    parser.add_argument("--lang", default="cn", choices=["cn", "sc", "en"], help="Feed language.")
    parser.add_argument("--output", default="", help="Output .xlsx path.")
    parser.add_argument(
        "--format",
        choices=["visible", "raw"],
        default="visible",
        help="visible maintains one web-style sheet; raw exports all JSON fields.",
    )
    parser.add_argument(
        "--types",
        nargs="*",
        default=None,
        help="Optional odds type codes to scrape. Default: all known page odds types.",
    )
    parser.add_argument("--proxy-host", default="", help="Optional HTTP/HTTPS proxy IP or host.")
    parser.add_argument("--proxy-port", default="", help="Optional HTTP/HTTPS proxy port.")
    parser.add_argument(
        "--no-verify-ssl",
        action="store_true",
        help="Disable HTTPS certificate verification. Use only with trusted proxies.",
    )
    return parser.parse_args(argv)


def run_scraper(args: argparse.Namespace) -> RunScrapeResult:
    proxy_host = getattr(args, "proxy_host", "")
    proxy_port = getattr(args, "proxy_port", "")
    current_time = require_run_allowed(proxy_host=proxy_host, proxy_port=proxy_port)
    print(f"[{current_time:%Y-%m-%d %H:%M:%S}] 开始更新澳彩全場三合一賠率", flush=True)
    timestamp = current_time.strftime("%Y%m%d_%H%M%S")
    if args.output:
        output = Path(args.output)
    elif args.format == "visible":
        output = app_base_dir() / "outputs" / "清洗后的表格.xlsx"
    else:
        output = app_base_dir() / "outputs" / f"macauslot_soccer_odds_raw_{timestamp}.xlsx"

    verify_ssl = not getattr(args, "no_verify_ssl", False)
    client = MacauslotClient(
        entry_url=args.url,
        proxy_host=proxy_host,
        proxy_port=proxy_port,
        verify_ssl=verify_ssl,
    )
    if client.proxy_url:
        print(f"使用网络代理：{client.proxy_url}", flush=True)
    if not client.verify_ssl:
        print("已关闭 SSL 证书校验，仅建议在可信代理环境下使用。", flush=True)
    # Prime cookies and solve any WAF challenge before the JSON feed requests.
    print("正在连接网页并处理访问校验...", flush=True)
    try:
        client.fetch_text(args.url)
    except urllib.error.URLError as exc:
        print(f"入口页面连接失败，将继续尝试 JSON 接口：{exc}", flush=True)

    if args.types:
        type_codes = args.types[:]
    elif args.format == "visible":
        type_codes = ["index"]
    else:
        type_codes = BASE_ODDS_TYPES[:]

    if args.types is None and args.format == "raw":
        for special in discover_special_types(client):
            if special not in type_codes:
                type_codes.append(special)

    results: list[ScrapeResult] = []
    for type_code in type_codes:
        print(f"正在抓取：{ODDS_TYPE_NAMES.get(type_code, type_code)}", flush=True)
        try:
            result = scrape_type(client, type_code, args.lang)
        except Exception as exc:
            result = ScrapeResult(
                type_code=type_code,
                type_name=ODDS_TYPE_NAMES.get(type_code, type_code),
                source="",
                event_rows=[],
                odds_rows=[],
                status="error",
                note=str(exc),
            )
        results.append(result)
        print(
            f"{result.type_code}: events={len(result.event_rows)}, rows={len(result.odds_rows)}, status={result.status}",
            flush=True,
        )

    if args.format == "raw":
        print("正在写入原始字段工作簿...", flush=True)
        build_workbook(results, output.resolve(), args.lang, args.url)
        updated_count, appended_count = 0, 0
    else:
        require_visible_scrape_data(results)
        print("正在按样表格式更新 Excel...", flush=True)
        matches = raw_to_sample_matches(results[0].odds_rows) if results else []
        updated_count, appended_count = update_sample_workbook(output.resolve(), matches)
        print(f"Updated cells: {updated_count}; appended matches: {appended_count}", flush=True)
    print(f"Saved: {output.resolve()}", flush=True)
    print(f"[{now():%Y-%m-%d %H:%M:%S}] 完成", flush=True)
    return RunScrapeResult(output.resolve(), updated_count=updated_count, appended_count=appended_count)


def main() -> None:
    args = parse_args()
    run_scraper(args)


if __name__ == "__main__":
    main()
