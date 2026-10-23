#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日美股财经简报 —— 抓取 + 排版 + 发送

数据源（均为公开 RSS / JSON 接口）
  - CNBC Top News      https://www.cnbc.com/id/100003114/device/rss/rss.html
  - CNBC Markets       https://www.cnbc.com/id/20910258/device/rss/rss.html
  - Yahoo Finance      https://feeds.finance.yahoo.com/rss/2.0/headline
  - MarketWatch        https://feeds.content.dowjones.io/public/rss/mw_topstories
行情
  - CNBC Quote API     quote.cnbc.com

用法
  python3 daily_brief.py --dry-run          # 只生成 HTML，不发信
  python3 daily_brief.py                    # 生成并发信
  python3 daily_brief.py --to a@b.com       # 临时指定收件人
"""

import argparse
import html
import json
import logging
import os
import re
import smtplib
import sys
import time
from datetime import datetime, timedelta, timezone
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr, parsedate_to_datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import requests

# Windows 命令行默认 GBK，强制 UTF-8 避免中文日志乱码
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

# ────────────────────────────── 配置 ──────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"

DEFAULT_CONFIG = {
    "recipients": ["Ricky.Jiang@yict.com.cn"],
    "smtp": {
        "host": "smtp.qq.com",
        "port": 465,
        "use_ssl": True,
        "user": "",
        "password": "",
        "sender_name": "每日财经简报"
    },
    "subject_prefix": "【美股晨报】",
    "max_items_per_source": 8,
    "total_items": 10,
    "hours_lookback": 30,
    "output_dir": "out"
}

# 行情标的
QUOTE_SYMBOLS = [
    (".SPX",   "S&P 500",   "index"),
    (".IXIC",  "纳斯达克",   "index"),
    (".DJI",   "道琼斯",     "index"),
    (".VIX",   "VIX 恐慌指数", "index"),
    ("US10Y",  "10 年期美债", "yield"),
    (".DXY",   "美元指数",   "index"),
    ("@GC.1",  "黄金",       "commodity"),
    ("@CL.1",  "WTI 原油",   "commodity"),
    ("@BTC.1", "比特币",     "crypto"),
]

FEEDS = [
    ("CNBC",        "https://www.cnbc.com/id/100003114/device/rss/rss.html"),
    ("CNBC 市场",   "https://www.cnbc.com/id/20910258/device/rss/rss.html"),
    ("Yahoo Finance", "https://feeds.finance.yahoo.com/rss/2.0/headline?s=%5EGSPC&region=US&lang=en-US"),
    ("MarketWatch", "https://feeds.content.dowjones.io/public/rss/mw_topstories"),
]

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36"),
    "Accept": "application/rss+xml, application/xml, application/json, text/xml, */*",
    "Accept-Language": "en-US,en;q=0.9",
}

# 个人理财 / 民生类关键词 —— 用于过滤噪音
NOISE_PATTERNS = [
    r"\bI'm \d+ and\b", r"\bmy (partner|husband|wife|mother|father|son|daughter)\b",
    r"social security", r"medicare", r"401\(k\)", r"\bIRA\b", r"roth",
    r"student loan", r"mortgage", r"credit card", r"life insurance",
    r"best .*(CD|savings|checking) rates", r"refinanc", r"annuit",
    r"renting", r"inheritance", r"estate plan", r"diabetes", r"health insurance",
    r"car insurance", r"pet insurance", r"personal loan", r"debt",
]

log = logging.getLogger("brief")


# ────────────────────────────── 工具函数 ──────────────────────────────
def load_config() -> dict:
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if CONFIG_PATH.exists():
        try:
            user_cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            for k, v in user_cfg.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as e:
            log.warning("config.json 解析失败，使用默认配置：%s", e)
    # 环境变量覆盖（便于在 CI / cron 中注入密钥）
    if os.getenv("SMTP_USER"):
        cfg["smtp"]["user"] = os.environ["SMTP_USER"]
    if os.getenv("SMTP_PASSWORD"):
        cfg["smtp"]["password"] = os.environ["SMTP_PASSWORD"]
    if os.getenv("BRIEF_TO"):
        cfg["recipients"] = [x.strip() for x in os.environ["BRIEF_TO"].split(",") if x.strip()]
    return cfg


def clean_text(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def is_noise(title: str) -> bool:
    t = title.lower()
    return any(re.search(p, t, re.I) for p in NOISE_PATTERNS)


def parse_feed(xml_text: str):
    """解析 RSS/Atom，返回 [(title, link, pub_dt_utc, summary, source_name)]"""
    out = []
    try:
        root = ET.fromstring(xml_text.encode("utf-8") if isinstance(xml_text, str) else xml_text)
    except ET.ParseError as e:
        log.warning("XML 解析失败：%s", e)
        return out

    def localname(tag):
        return tag.split("}")[-1].lower()

    for node in root.iter():
        if localname(node.tag) not in ("item", "entry"):
            continue
        title = link = summary = ""
        pub = None
        for ch in node:
            ln = localname(ch.tag)
            if ln == "title":
                title = clean_text(ch.text or "")
            elif ln == "link":
                link = (ch.text or "").strip() or ch.attrib.get("href", "").strip()
            elif ln in ("description", "summary", "content"):
                summary = clean_text(ch.text or "")
            elif ln in ("pubdate", "published", "updated", "dc:date"):
                raw = clean_text(ch.text or "")
                if raw:
                    try:
                        pub = parsedate_to_datetime(raw)
                    except Exception:
                        try:
                            pub = datetime.fromisoformat(raw.replace("Z", "+00:00"))
                        except Exception:
                            pub = None
        if pub and pub.tzinfo is None:
            pub = pub.replace(tzinfo=timezone.utc)
        if title and link:
            out.append((title, link, pub, summary))
    return out


def fetch_feed(name: str, url: str, retries: int = 2):
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 200:
                items = parse_feed(r.text)
                log.info("  %-14s %d 条", name, len(items))
                return [(t, l, p, s, name) for (t, l, p, s) in items]
            log.warning("  %-14s HTTP %s", name, r.status_code)
        except Exception as e:
            log.warning("  %-14s 失败(%s/%s)：%s", name, attempt + 1, retries + 1, type(e).__name__)
        time.sleep(1.5 * (attempt + 1))
    return []


def fetch_quotes():
    symbols = "|".join(s for s, _, _ in QUOTE_SYMBOLS)
    url = ("https://quote.cnbc.com/quote-html-webservice/restQuote/symbolType/symbol"
           f"?symbols={symbols}&requestMethod=itv&noform=1&partnerId=2&fund=1&exthrs=1"
           "&output=json&events=1")
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        raw = r.json()["FormattedQuoteResult"]["FormattedQuote"]
        if isinstance(raw, dict):
            raw = [raw]
    except Exception as e:
        log.error("  行情获取失败：%s", e)
        return []

    by_symbol = {q.get("symbol"): q for q in raw}
    rows = []
    for sym, label, kind in QUOTE_SYMBOLS:
        q = by_symbol.get(sym)
        if not q:
            continue
        try:
            last = float(str(q.get("last", "")).replace(",", "").replace("%", ""))
        except ValueError:
            continue
        pct = str(q.get("change_pct", "")).replace("%", "").strip()
        try:
            pct = float(pct)
        except ValueError:
            pct = None
        chg = str(q.get("change", "")).strip()
        rows.append({
            "label": label,
            "kind": kind,
            "last": last,
            "display": q.get("last", ""),
            "change": chg,
            "pct": pct,
            "is_pct_instrument": str(q.get("last", "")).endswith("%"),
        })
    log.info("  行情 %d 项", len(rows))
    return rows


# ────────────────────────────── 选题逻辑 ──────────────────────────────
PRIORITY_WORDS = [
    ("fed", 10), ("powell", 9), ("inflation", 9), ("cpi", 9), ("jobs report", 8),
    ("rate cut", 10), ("rates", 6), ("tariff", 9), ("treasury", 7), ("yield", 7),
    ("earnings", 8), ("revenue", 6), ("guidance", 7), ("nvidia", 8), ("apple", 7),
    ("tesla", 7), ("microsoft", 7), ("meta", 6), ("amazon", 6), ("alphabet", 6),
    ("openai", 7), ("ai ", 6), ("semiconductor", 7), ("chip", 6), ("oil", 6),
    ("opec", 7), ("gold", 5), ("bitcoin", 7), ("crypto", 5), ("sec ", 5),
    ("merger", 6), ("acquisition", 6), ("layoff", 5), ("ipo", 6), ("gdp", 8),
    ("recession", 8), ("selloff", 7), ("rally", 6), ("dollar", 5), ("china", 6),
    ("nasdaq", 6), ("s&p", 6), ("dow", 5), ("futures", 6), ("bank", 5),
]


def score_item(title: str, pub, now) -> float:
    score = 0.0
    t = " " + title.lower() + " "
    for w, s in PRIORITY_WORDS:
        if w in t:
            score += s
    if pub:
        age_h = (now - pub).total_seconds() / 3600
        if age_h <= 3:
            score += 8
        elif age_h <= 8:
            score += 5
        elif age_h <= 16:
            score += 2
        elif age_h > 48:
            score -= 10
    if len(title) < 25:
        score -= 3
    return score


def select_items(all_items, cfg):
    now = datetime.now(timezone.utc)
    cutoff = now - timedelta(hours=cfg["hours_lookback"])

    seen_titles, seen_links = set(), set()
    candidates = []
    for title, link, pub, summary, source in all_items:
        if is_noise(title):
            continue
        if pub and pub < cutoff:
            continue
        key = re.sub(r"[^a-z0-9]", "", title.lower())[:60]
        if key in seen_titles or link in seen_links:
            continue
        seen_titles.add(key)
        seen_links.add(link)
        candidates.append({
            "title": title, "link": link, "pub": pub,
            "summary": summary, "source": source,
            "score": score_item(title, pub, now),
        })

    candidates.sort(key=lambda x: x["score"], reverse=True)
    log.info("  去重后候选 %d 条，取前 %d 条", len(candidates), cfg["total_items"])
    return candidates[: cfg["total_items"]]


# ────────────────────────────── 中文要点 ──────────────────────────────
# 轻量关键词映射：把英文标题转成中文导语。
# 没有调用大模型，纯规则；如接入 LLM 可替换此函数。
GLOSSARY = {
    r"\bfed\b|federal reserve": "美联储",
    r"\brate cut|cut rates|easing": "降息",
    r"\brate hike|raise rates": "加息",
    r"inflation": "通胀",
    r"cpi|consumer price": "CPI 物价数据",
    r"jobs report|payrolls|nonfarm": "就业数据",
    r"earnings|results": "财报",
    r"guidance|outlook": "业绩指引",
    r"revenue": "营收",
    r"tariff": "关税",
    r"treasury|yield": "美债收益率",
    r"stocks?|equit(y|ies)": "股市",
    r"futures": "股指期货",
    r"rally|surge|jump|soar|climb": "上涨",
    r"slide|fall|drop|plunge|tumble|selloff|slump": "下跌",
    r"nasdaq": "纳斯达克",
    r"s&p 500": "标普 500",
    r"dow": "道琼斯",
    r"nvidia": "英伟达",
    r"apple": "苹果",
    r"tesla": "特斯拉",
    r"microsoft": "微软",
    r"meta": "Meta",
    r"amazon": "亚马逊",
    r"alphabet|google": "谷歌",
    r"openai": "OpenAI",
    r"bitcoin|crypto": "加密货币",
    r"\bai\b|artificial intelligence": "人工智能",
    r"semiconductor|chip": "芯片半导体",
    r"oil|crude": "原油",
    r"gold": "黄金",
    r"dollar": "美元",
    r"china|chinese": "中国",
    r"merger|acquisition|deal": "并购交易",
    r"ipo": "IPO",
    r"layoff|job cuts": "裁员",
    r"gdp": "GDP",
    r"recession": "衰退",
    r"bank": "银行",
    r"housing|home": "房地产",
    r"trade": "贸易",
}


def cn_hint(title: str) -> str:
    """从英文标题提取中文关键词标签，作为快速扫读的辅助。"""
    tags = []
    tl = title.lower()
    for pat, cn in GLOSSARY.items():
        if re.search(pat, tl) and cn not in tags:
            tags.append(cn)
    return " · ".join(tags[:4])


# ────────────────────────────── HTML 渲染 ──────────────────────────────
def fmt_pct(pct):
    if pct is None:
        return "—"
    sign = "+" if pct > 0 else ""
    return f"{sign}{pct:.2f}%"


def color_of(pct):
    """中国市场惯例：涨红跌绿。"""
    if pct is None:
        return "#6b7280"
    if pct > 0:
        return "#c62828"   # 红
    if pct < 0:
        return "#2e7d32"   # 绿
    return "#6b7280"


def build_quote_cards(quotes):
    if not quotes:
        return '<p style="color:#888;font-size:13px;">行情数据暂时不可用。</p>'
    cards = []
    for q in quotes:
        pct = q["pct"]
        c = color_of(pct)
        arrow = "▲" if (pct or 0) > 0 else ("▼" if (pct or 0) < 0 else "▪")
        chg_txt = q["change"] if q["change"] else "—"
        if q["is_pct_instrument"]:
            chg_txt = ""
        cards.append(f"""
        <td style="padding:7px 10px;border-bottom:1px solid #eceff1;">
          <div style="font-size:12px;color:#78909c;letter-spacing:.3px;">{q['label']}</div>
          <div style="font-size:17px;font-weight:600;color:#1a237e;line-height:1.4;">{q['display']}</div>
          <div style="font-size:12px;color:{c};font-weight:600;">{arrow} {chg_txt} ({fmt_pct(pct)})</div>
        </td>""")

    rows, per_row = [], 3
    for i in range(0, len(cards), per_row):
        chunk = cards[i:i + per_row]
        while len(chunk) < per_row:
            chunk.append('<td style="padding:7px 10px;"></td>')
        rows.append("<tr>" + "".join(chunk) + "</tr>")
    return f"""
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"
           style="border:1px solid #e0e6ea;border-radius:6px;border-collapse:separate;overflow:hidden;background:#fff;">
      {''.join(rows)}
    </table>"""


def build_items_html(items):
    if not items:
        return '<p style="color:#888;">今日无符合条件的新闻。</p>'
    blocks = []
    for i, it in enumerate(items, 1):
        hint = cn_hint(it["title"])
        when = ""
        if it["pub"]:
            bj = it["pub"].astimezone(timezone(timedelta(hours=8)))
            when = bj.strftime("%m-%d %H:%M")
        summary = it["summary"][:180]
        if summary and len(it["summary"]) > 180:
            summary += "…"
        hint_html = (f'<div style="font-size:12px;color:#00695c;background:#e0f2f1;'
                     f'display:inline-block;padding:2px 8px;border-radius:10px;margin:4px 0 6px;">'
                     f'{hint}</div>') if hint else ""
        sum_html = (f'<div style="font-size:13px;color:#546e7a;line-height:1.6;margin-top:4px;">'
                    f'{html.escape(summary)}</div>') if summary else ""
        blocks.append(f"""
        <div style="padding:14px 0;border-bottom:1px solid #eceff1;">
          <div style="font-size:11px;color:#90a4ae;margin-bottom:4px;">
            <span style="background:#eceff1;padding:1px 7px;border-radius:3px;">{it['source']}</span>
            &nbsp;{when}
          </div>
          <a href="{it['link']}" style="font-size:15px;font-weight:600;color:#0d47a1;
             text-decoration:none;line-height:1.45;">{i}. {html.escape(it['title'])}</a>
          {hint_html}
          {sum_html}
        </div>""")
    return "".join(blocks)


def build_html(quotes, items, config):
    now_bj = datetime.now(timezone(timedelta(hours=8)))
    date_cn = now_bj.strftime("%Y年%m月%d日")
    weekday_cn = "一二三四五六日"[now_bj.weekday()]

    up = sum(1 for q in quotes if (q["pct"] or 0) > 0)
    down = sum(1 for q in quotes if (q["pct"] or 0) < 0)
    mood = "普涨" if down == 0 and up else ("普跌" if up == 0 and down else "涨跌互现")

    spx = next((q for q in quotes if q["label"] == "S&P 500"), None)
    ixic = next((q for q in quotes if q["label"] == "纳斯达克"), None)
    headline_bits = []
    if spx and spx["pct"] is not None:
        headline_bits.append(f"标普 500 {fmt_pct(spx['pct'])}")
    if ixic and ixic["pct"] is not None:
        headline_bits.append(f"纳指 {fmt_pct(ixic['pct'])}")
    headline = "，".join(headline_bits) if headline_bits else "行情待更新"

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>美股晨报 {date_cn}</title></head>
<body style="margin:0;padding:0;background:#f5f7f9;font-family:-apple-system,'Segoe UI','Microsoft YaHei',Arial,sans-serif;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f5f7f9;padding:20px 10px;">
<tr><td align="center">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"
       style="max-width:660px;background:#ffffff;border-radius:8px;overflow:hidden;
              box-shadow:0 1px 4px rgba(0,0,0,.06);">

  <!-- 页头 -->
  <tr><td style="background:linear-gradient(135deg,#0d47a1,#1565c0);padding:22px 26px;">
    <div style="color:#90caf9;font-size:11px;letter-spacing:2px;margin-bottom:6px;">DAILY MARKET BRIEF</div>
    <div style="color:#ffffff;font-size:21px;font-weight:600;letter-spacing:.5px;">美股晨报</div>
    <div style="color:#bbdefb;font-size:13px;margin-top:6px;">{date_cn} 星期{weekday_cn} · 北京时间 {now_bj.strftime('%H:%M')}</div>
  </td></tr>

  <!-- 一句话概览 -->
  <tr><td style="padding:18px 26px 4px;">
    <div style="font-size:14px;color:#37474f;line-height:1.7;">
      <strong style="color:#0d47a1;">今日概览：</strong>{headline}；主要资产{mood}（{up} 涨 / {down} 跌）。
    </div>
  </td></tr>

  <!-- 行情 -->
  <tr><td style="padding:14px 26px 6px;">
    <div style="font-size:13px;font-weight:700;color:#0d47a1;letter-spacing:1px;
                border-left:3px solid #1565c0;padding-left:8px;margin-bottom:10px;">市场行情</div>
    {build_quote_cards(quotes)}
    <div style="font-size:11px;color:#b0bec5;margin-top:6px;">数据来源：CNBC Quote API · 涨红跌绿</div>
  </td></tr>

  <!-- 要闻 -->
  <tr><td style="padding:16px 26px 8px;">
    <div style="font-size:13px;font-weight:700;color:#0d47a1;letter-spacing:1px;
                border-left:3px solid #1565c0;padding-left:8px;">今日要闻</div>
    {build_items_html(items)}
  </td></tr>

  <!-- 页脚 -->
  <tr><td style="padding:18px 26px 22px;background:#fafbfc;border-top:1px solid #eceff1;">
    <div style="font-size:11px;color:#90a4ae;line-height:1.8;">
      本邮件由脚本自动生成，信息源为 CNBC、Yahoo Finance、MarketWatch 的公开 RSS 接口。<br>
      中文标签由关键词规则自动匹配，仅供快速扫读，请以原文链接内容为准。<br>
      本简报不构成任何投资建议。
    </div>
  </td></tr>

</table>
</td></tr></table>
</body></html>"""


def build_plain(quotes, items, config):
    now_bj = datetime.now(timezone(timedelta(hours=8)))
    lines = [f"美股晨报 · {now_bj.strftime('%Y-%m-%d %H:%M')}（北京时间）", "=" * 46, "", "【市场行情】"]
    for q in quotes:
        lines.append(f"  {q['label']:<14} {q['display']:>12}  {fmt_pct(q['pct'])}")
    lines += ["", "【今日要闻】", ""]
    for i, it in enumerate(items, 1):
        hint = cn_hint(it["title"])
        lines.append(f"{i}. {it['title']}")
        if hint:
            lines.append(f"   [{hint}]")
        lines.append(f"   {it['link']}")
        lines.append("")
    lines += ["-" * 46,
              "信息源：CNBC / Yahoo Finance / MarketWatch 公开 RSS",
              "本简报不构成投资建议。"]
    return "\n".join(lines)


# ────────────────────────────── 发送 ──────────────────────────────
def send_mail(subject, html_body, text_body, config):
    smtp_cfg = config["smtp"]
    user = smtp_cfg.get("user") or ""
    pwd = smtp_cfg.get("password") or ""
    if not user or not pwd:
        raise RuntimeError(
            "未配置发件邮箱。请在 config.json 的 smtp.user / smtp.password 中填写，"
            "或设置环境变量 SMTP_USER 与 SMTP_PASSWORD。"
        )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = formataddr((str(Header(smtp_cfg.get("sender_name", "每日财经简报"), "utf-8")), user))
    msg["To"] = ", ".join(config["recipients"])
    msg.attach(MIMEText(text_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    host, port = smtp_cfg["host"], int(smtp_cfg["port"])
    if smtp_cfg.get("use_ssl", True):
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        server.starttls()
    try:
        server.login(user, pwd)
        server.sendmail(user, config["recipients"], msg.as_string())
    finally:
        server.quit()
    log.info("已发送至：%s", ", ".join(config["recipients"]))


# ────────────────────────────── 主流程 ──────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="生成并发送美股每日简报")
    ap.add_argument("--dry-run", action="store_true", help="只生成 HTML，不发送邮件")
    ap.add_argument("--to", default="", help="临时覆盖收件人，多个用逗号分隔")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        datefmt="%H:%M:%S",
    )

    config = load_config()
    if args.to:
        config["recipients"] = [x.strip() for x in args.to.split(",") if x.strip()]

    log.info("① 抓取行情 …")
    quotes = fetch_quotes()

    log.info("② 抓取新闻源 …")
    all_items = []
    for name, url in FEEDS:
        all_items.extend(fetch_feed(name, url))

    log.info("③ 筛选排序 …")
    items = select_items(all_items, config)

    log.info("④ 渲染邮件 …")
    subject = f"{config['subject_prefix']}{datetime.now(timezone(timedelta(hours=8))).strftime('%Y-%m-%d')} 美股每日更新"
    html_body = build_html(quotes, items, config)
    text_body = build_plain(quotes, items, config)

    out_dir = BASE_DIR / config.get("output_dir", "out")
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d")
    html_path = out_dir / f"brief-{stamp}.html"
    html_path.write_text(html_body, encoding="utf-8")
    log.info("   已保存：%s", html_path)

    if args.dry_run:
        log.info("dry-run 模式，未发送邮件。")
        return 0

    log.info("⑤ 发送邮件 …")
    try:
        send_mail(subject, html_body, text_body, config)
    except Exception as e:
        log.error("发送失败：%s", e)
        return 2
    return 0


if __name__ == "__main__":
    sys.exit(main())
