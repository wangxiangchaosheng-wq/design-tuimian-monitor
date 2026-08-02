#!/usr/bin/env python3
"""2027届985/211高校设计类预推免近实时监测器。

数据原则：
1. 仅抓配置中的高校官方站点；
2. 识别接收推免/预报名通知，排除校内推免资格、拟推荐名单等信息；
3. 尽量按官网原有条目提取专业方向、考核方式和申请材料；
4. 自动生成分组、编号清单式 CSV / JSON / HTML，便于后续重建 Excel。
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import os
import re
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta, timezone
from html import escape
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin, urlparse, urldefrag
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from dateutil import parser as date_parser
from pypdf import PdfReader

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config" / "schools.json"
DETAILED_SEED_PATH = BASE_DIR / "data" / "notices_seed_detailed.json"
STATE_PATH = BASE_DIR / "data" / "notices.json"
CSV_PATH = BASE_DIR / "data" / "notices.csv"
HTML_PATH = BASE_DIR / "docs" / "index.html"
CHANGES_PATH = BASE_DIR / "data" / "changes.json"
TZ_CN = timezone(timedelta(hours=8))
TARGET_YEAR = 2027

INCOMING_TERMS = [
    "接收推荐免试", "接收外校", "招收推荐免试", "推荐免试研究生预报名", "推免生预报名",
    "推免预报名", "预推免", "预选拔", "免试攻读研究生", "接收优秀应届本科毕业生免试",
    "接收推免", "推荐免试攻读", "含直接攻博"
]
NEGATIVE_TERMS = [
    "推荐2027届", "推荐免试资格", "推免资格", "推荐资格", "校内推免", "推免工作细则",
    "拟推荐名单", "综合排名", "推荐名额", "候选人名单", "获得推免资格", "推免生资格审核"
]
DESIGN_TERMS = [
    "设计学", "设计", "艺术设计", "工业设计", "交互设计", "智能交互", "信息艺术设计",
    "数字媒体艺术", "数字媒体技术", "新媒体艺术", "视觉传达", "环境设计", "环境艺术设计",
    "产品设计", "服务设计", "媒体与传达", "人工智能与数据设计", "互动媒体设计与技术",
    "互联网+创新设计", "未来人居", "智慧创新设计", "创意与创新", "建筑学", "建筑",
    "城乡规划", "风景园林", "景观", "城市设计", "艺术与科技", "智能游戏交互"
]
DATE_RE = re.compile(r"(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日?")
TIME_RE = re.compile(r"(\d{1,2})\s*[:：]\s*(\d{2})")
ITEM_PREFIX_RE = re.compile(r"^\s*(?:[（(]?[一二三四五六七八九十百]+[）)、.]|\d{1,2}[.、）)]|[①②③④⑤⑥⑦⑧⑨⑩]|[-•·])\s*")
APPLICATION_DOMAIN_HINTS = (
    "yzbm", "yjszs", "ssyjsbm", "gsas", "sszs", "yzsbm", "yzfs", "yjsfs", "graduate-admission"
)


def now_cn() -> datetime:
    return datetime.now(TZ_CN)


def fmt_dt(dt: datetime | None) -> str:
    return dt.astimezone(TZ_CN).strftime("%Y-%m-%d %H:%M") if dt else ""


def normalize_space(text: str) -> str:
    return re.sub(r"[ \t\u3000]+", " ", text or "").strip()


def normalize_url(url: str) -> str:
    return urldefrag((url or "").strip())[0]


def clean_item(text: str) -> str:
    text = ITEM_PREFIX_RE.sub("", normalize_space(text))
    return text.strip("；;。 ")


def numbered(items: Iterable[str] | str) -> str:
    if isinstance(items, str):
        values = [clean_item(x) for x in items.splitlines() if clean_item(x)]
    else:
        values = [clean_item(str(x)) for x in items if clean_item(str(x))]
    return "\n".join(f"{i}. {v}" for i, v in enumerate(values, 1))


def plain_items(value: str) -> list[str]:
    return [clean_item(x) for x in (value or "").splitlines() if clean_item(x)]


@dataclass
class Notice:
    tier: str
    school: str
    unit: str
    title: str
    programs: str
    published_at: str
    application_start: str
    deadline: str
    status: str
    assessment: str
    source_url: str
    application_url: str
    materials: str
    cross_major: str
    confidence: str
    first_seen_at: str
    last_checked_at: str
    notes: str
    content_hash: str = ""
    manual_review: bool = False
    score: int = 0

    def key(self) -> str:
        token = f"{self.school}|{self.unit}|{normalize_url(self.source_url)}"
        return hashlib.sha1(token.encode("utf-8")).hexdigest()


def host_is_official(candidate: str, configured_url: str) -> bool:
    c = (urlparse(candidate).hostname or "").lower()
    root = (urlparse(configured_url).hostname or "").lower()
    if not c or not root:
        return False
    return c == root or c.endswith("." + root) or root.endswith("." + c)


def parse_cn_date(text: str) -> datetime | None:
    m = DATE_RE.search(text or "")
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    hour = minute = 0
    tail = (text or "")[m.end():m.end() + 30]
    tm = TIME_RE.search(tail)
    if tm:
        hour, minute = map(int, tm.groups())
        if hour == 24:
            hour = 0
            return datetime(y, mo, d, hour, minute, tzinfo=TZ_CN) + timedelta(days=1)
    elif "下午" in tail:
        hour = 15
    elif "上午" in tail:
        hour = 10
    return datetime(y, mo, d, hour, minute, tzinfo=TZ_CN)


def extract_date_near(text: str, keywords: list[str], prefer: str) -> str:
    found: list[datetime] = []
    for kw in keywords:
        pos = 0
        while True:
            idx = text.find(kw, pos)
            if idx < 0:
                break
            snippet = text[max(0, idx - 40):idx + 180]
            for m in DATE_RE.finditer(snippet):
                dt = parse_cn_date(snippet[m.start():m.end() + 25])
                if dt and dt.year in {TARGET_YEAR - 1, TARGET_YEAR}:
                    found.append(dt)
            pos = idx + len(kw)
    if not found:
        return ""
    dt = min(found) if prefer == "min" else max(found)
    return dt.strftime("%Y-%m-%d %H:%M")


def extract_published(text: str, title: str) -> str:
    value = extract_date_near(text[:5000], ["发布时间", "发布日期", "发布于", "日期"], "min")
    if value:
        return value[:10]
    dt = parse_cn_date(title)
    return dt.strftime("%Y-%m-%d") if dt else ""


def compute_status(deadline: str) -> str:
    if not deadline:
        return "截止时间待核验"
    try:
        dt = date_parser.parse(deadline)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=TZ_CN)
    except Exception:
        return "截止时间待核验"
    delta = dt - now_cn()
    if delta.total_seconds() < 0:
        return "已截止"
    if delta <= timedelta(days=3):
        return "3天内截止"
    return "报名中"


def classify(title: str, text: str) -> tuple[int, int, bool]:
    sample = normalize_space(title + " " + text[:30000])
    incoming = sum(4 if t in title else 1 for t in INCOMING_TERMS if t in sample)
    negative = sum(5 if t in title else 2 for t in NEGATIVE_TERMS if t in sample)
    design = sum(3 if t in title else 1 for t in DESIGN_TERMS if t in sample)
    score = incoming + design + (3 if str(TARGET_YEAR) in sample else -4) - negative
    return score, design, negative >= incoming + 3


def unique_ordered(items: Iterable[str], limit: int = 30) -> list[str]:
    out = []
    seen = set()
    for x in items:
        x = clean_item(x)
        if not x or len(x) < 2 or x in seen:
            continue
        seen.add(x)
        out.append(x)
        if len(out) >= limit:
            break
    return out


def lines_from_text(text: str) -> list[str]:
    return [normalize_space(x) for x in re.split(r"[\r\n]+", text or "") if normalize_space(x)]


def extract_section_items(text: str, headings: list[str], stop_headings: list[str], max_items: int = 20) -> list[str]:
    lines = lines_from_text(text)
    start = None
    for i, line in enumerate(lines):
        if any(h in line for h in headings):
            start = i + 1
            # The heading line may contain the first item after a colon.
            if "：" in line or ":" in line:
                tail = re.split(r"[：:]", line, maxsplit=1)[1]
                if clean_item(tail):
                    lines.insert(start, tail)
            break
    if start is None:
        return []
    out = []
    for line in lines[start:start + 80]:
        if out and any(h in line for h in stop_headings) and len(line) < 40:
            break
        if len(line) > 500:
            continue
        if ITEM_PREFIX_RE.match(line) or len(line) <= 180:
            val = clean_item(line)
            if val and not any(val == h for h in headings):
                out.append(val)
        if len(out) >= max_items:
            break
    return unique_ordered(out, max_items)


def extract_programs(text: str) -> list[str]:
    section = extract_section_items(
        text,
        ["招生专业", "招生方向", "专业方向", "申请方向", "招收专业", "招生学科"],
        ["申请条件", "报名时间", "申请材料", "报名材料", "考核", "复试"],
        20,
    )
    if section:
        return section
    found = [t for t in DESIGN_TERMS if t in text]
    found.sort(key=len, reverse=True)
    return unique_ordered(found, 15)


def extract_assessment(text: str) -> list[str]:
    return extract_section_items(
        text,
        ["考核方式", "考核安排", "复试安排", "选拔程序", "审核与考核", "考核录取"],
        ["申请材料", "报名材料", "联系方式", "其他事项", "咨询"],
        16,
    )


def extract_materials(text: str) -> list[str]:
    items = extract_section_items(
        text,
        ["申请材料", "报名材料", "提交材料", "材料要求", "上传材料"],
        ["考核方式", "复试安排", "选拔程序", "联系方式", "其他事项", "注意事项"],
        30,
    )
    if items:
        return items
    # Conservative fallback: only collect lines explicitly mentioning common documents.
    keys = ["身份证", "学生证", "学籍", "成绩单", "排名", "外语", "推荐信", "作品集", "个人陈述", "研究计划", "申请表", "承诺书", "获奖", "论文", "专利"]
    return unique_ordered([line for line in lines_from_text(text) if any(k in line for k in keys) and len(line) <= 220], 20)


def extract_cross_major(text: str) -> str:
    matches = []
    for line in lines_from_text(text):
        if any(h in line for h in ["专业不限", "跨专业", "非艺术", "欢迎", "相关专业", "专业背景"]):
            if len(line) <= 220:
                matches.append(line)
    return "；".join(unique_ordered(matches, 4)) or "待核验"


def detect_application_url(links: list[tuple[str, str]], source_url: str) -> str:
    for href, label in links:
        token = (href + " " + label).lower()
        if any(h in token for h in APPLICATION_DOMAIN_HINTS) or any(k in label for k in ["报名系统", "申请系统", "预报名系统", "研究生招生系统"]):
            return href
    return ""


def html_to_text(content: bytes, url: str) -> tuple[str, str, list[tuple[str, str]]]:
    soup = BeautifulSoup(content, "html.parser")
    for tag in soup(["script", "style", "noscript", "svg"]):
        tag.decompose()
    title_node = soup.find("h1") or soup.find("title") or soup
    title = normalize_space(title_node.get_text(" ", strip=True))[:220]
    # Preserve line structure; this is essential for numbered-material extraction.
    text = soup.get_text("\n", strip=True)
    text = re.sub(r"\n{3,}", "\n\n", text)
    links: list[tuple[str, str]] = []
    for a in soup.find_all("a", href=True):
        href = normalize_url(urljoin(url, a["href"]))
        label = normalize_space(a.get_text(" ", strip=True))
        if href.startswith("http"):
            links.append((href, label))
    return title, text, list(dict.fromkeys(links))


def pdf_to_text(content: bytes, url: str) -> tuple[str, str, list[tuple[str, str]]]:
    reader = PdfReader(io.BytesIO(content))
    text = "\n".join(page.extract_text() or "" for page in reader.pages[:30])
    text = re.sub(r"\n{3,}", "\n\n", text)
    title = normalize_space(text[:200]) or Path(urlparse(url).path).name
    return title, text, []


class Crawler:
    def __init__(self, timeout: int = 20, delay: float = 1.3, max_pages_per_source: int = 20):
        self.timeout = timeout
        self.delay = delay
        self.max_pages = max_pages_per_source
        self.session = requests.Session()
        self.robots: dict[str, RobotFileParser | None] = {}
        self.session.headers.update({
            "User-Agent": "DesignPushMonitor/2.0 (+academic-information-monitor; low-frequency; official-sites-only)",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.5",
        })

    def allowed_by_robots(self, url: str) -> bool:
        p = urlparse(url)
        host = p.netloc.lower()
        if host not in self.robots:
            rp = RobotFileParser()
            rp.set_url(f"{p.scheme}://{host}/robots.txt")
            try:
                r = self.session.get(rp.url, timeout=min(self.timeout, 8))
                if r.ok:
                    rp.parse(r.text.splitlines())
                    self.robots[host] = rp
                else:
                    self.robots[host] = None
            except requests.RequestException:
                self.robots[host] = None
        rp = self.robots[host]
        return True if rp is None else rp.can_fetch(self.session.headers.get("User-Agent", "*"), url)

    def fetch(self, url: str) -> tuple[bytes, str] | None:
        if not self.allowed_by_robots(url):
            print(f"[INFO] robots.txt disallows: {url}", file=sys.stderr)
            return None
        try:
            r = self.session.get(url, timeout=self.timeout, allow_redirects=True)
            r.raise_for_status()
            if len(r.content) > 30 * 1024 * 1024:
                return None
            return r.content, r.headers.get("Content-Type", "").lower()
        except requests.RequestException as exc:
            print(f"[WARN] fetch failed: {url} :: {exc}", file=sys.stderr)
            return None

    def crawl_source(self, school: dict, start_url: str) -> list[Notice]:
        queue = [(start_url, 0)]
        visited = set()
        notices = []
        while queue and len(visited) < self.max_pages:
            url, depth = queue.pop(0)
            url = normalize_url(url)
            if url in visited or not host_is_official(url, start_url):
                continue
            visited.add(url)
            result = self.fetch(url)
            time.sleep(self.delay)
            if not result:
                continue
            content, ctype = result
            try:
                title, text, links = pdf_to_text(content, url) if "pdf" in ctype or url.lower().endswith(".pdf") else html_to_text(content, url)
            except Exception as exc:
                print(f"[WARN] parse failed: {url} :: {exc}", file=sys.stderr)
                continue
            score, design_score, is_negative = classify(title, text)
            if score >= 7 and design_score >= 1 and not is_negative:
                checked = fmt_dt(now_cn())
                deadline = extract_date_near(text, ["截止", "系统关闭", "之前完成", "报名时间", "申请时间"], "max")
                start = extract_date_near(text, ["报名时间", "申请时间", "系统开放", "开始"], "min")
                programs = extract_programs(text)
                assessment = extract_assessment(text)
                materials = extract_materials(text)
                manual = not deadline or not materials
                notices.append(Notice(
                    tier=school.get("tier", "待核验"),
                    school=school["school"],
                    unit=self._guess_unit(school.get("units", ""), title, text),
                    title=title[:180],
                    programs=numbered(programs or [school.get("units", "相关方向")]),
                    published_at=extract_published(text, title),
                    application_start=start,
                    deadline=deadline,
                    status=compute_status(deadline),
                    assessment=numbered(assessment or ["考核方式以学院后续通知为准"]),
                    source_url=url,
                    application_url=detect_application_url(links, url),
                    materials=numbered(materials or ["申请材料清单待官网正文进一步核验"]),
                    cross_major=extract_cross_major(text),
                    confidence="官网自动采集" if not manual else "官网自动采集；关键字段待人工核验",
                    first_seen_at=checked,
                    last_checked_at=checked,
                    notes="自动识别；申请前必须打开官网逐项复核",
                    content_hash=hashlib.sha256(text.encode("utf-8")).hexdigest(),
                    manual_review=manual,
                    score=score,
                ))
            if depth < 1:
                for link, label in links:
                    if host_is_official(link, start_url) and any(k in (link + label) for k in [str(TARGET_YEAR), "推免", "推荐免试", "预报名"]):
                        queue.append((link, depth + 1))
        return notices

    @staticmethod
    def _guess_unit(units: str, title: str, text: str) -> str:
        for unit in [u.strip() for u in units.split("；") if u.strip()]:
            if unit in title or unit in text[:3000]:
                return unit
        return units.split("；")[0] if units else "相关学院"


def read_seed() -> list[Notice]:
    if not DETAILED_SEED_PATH.exists():
        return []
    payload = json.loads(DETAILED_SEED_PATH.read_text("utf-8"))
    checked = payload.get("generated_at") or fmt_dt(now_cn())
    rows = []
    for x in payload.get("notices", []):
        rows.append(Notice(
            tier=x.get("tier", "待核验"), school=x.get("school", ""), unit=x.get("unit", ""),
            title=x.get("title", ""), programs=numbered(x.get("programs", [])),
            published_at=x.get("published_at", ""), application_start=x.get("application_start", ""),
            deadline=x.get("deadline", ""), status=compute_status(x.get("deadline", "")),
            assessment=numbered(x.get("assessment", [])), source_url=x.get("source_url", ""),
            application_url=x.get("application_url", ""), materials=numbered(x.get("materials", [])),
            cross_major=x.get("cross_major", "待核验"), confidence=x.get("confidence", "官网已核验"),
            first_seen_at=checked, last_checked_at=checked, notes=x.get("notes", ""),
            content_hash=hashlib.sha256(json.dumps(x, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest(),
            manual_review="待" in x.get("confidence", ""), score=99,
        ))
    return rows


def load_state() -> dict[str, dict]:
    if not STATE_PATH.exists():
        return {}
    try:
        data = json.loads(STATE_PATH.read_text("utf-8"))
        return {x["key"]: x for x in data.get("notices", [])}
    except Exception:
        return {}


def merge_notices(seed: list[Notice], crawled: list[Notice]) -> tuple[list[Notice], list[dict]]:
    old = load_state()
    merged: dict[str, Notice] = {}
    changes = []
    for n in seed + crawled:
        key = n.key()
        prior = old.get(key)
        if prior:
            n.first_seen_at = prior.get("first_seen_at") or n.first_seen_at
            if n.content_hash and n.content_hash != prior.get("content_hash"):
                changes.append({"type": "updated", "tier": n.tier, "school": n.school, "unit": n.unit, "title": n.title, "url": n.source_url})
        else:
            changes.append({"type": "new", "tier": n.tier, "school": n.school, "unit": n.unit, "title": n.title, "url": n.source_url})
        # Seed/manual verified rows win; otherwise prefer the row with more extracted items.
        completeness = len(n.programs) + len(n.assessment) + len(n.materials) + (10000 if "全文已核验" in n.confidence else 0)
        if key not in merged:
            merged[key] = n
        else:
            prior_n = merged[key]
            prior_score = len(prior_n.programs) + len(prior_n.assessment) + len(prior_n.materials) + (10000 if "全文已核验" in prior_n.confidence else 0)
            if completeness > prior_score:
                merged[key] = n
    result = list(merged.values())
    for n in result:
        n.status = compute_status(n.deadline)
        n.last_checked_at = fmt_dt(now_cn())
    # UX ordering: tier -> school -> deadline. Same school is always contiguous.
    result.sort(key=lambda x: (0 if x.tier == "985" else 1, x.school, x.deadline or "9999-99-99", x.unit))
    return result, changes


def save_outputs(notices: list[Notice], changes: list[dict], schools: list[dict]) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {"generated_at": fmt_dt(now_cn()), "target_year": TARGET_YEAR, "notices": [{"key": n.key(), **asdict(n)} for n in notices]}
    STATE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")
    CHANGES_PATH.write_text(json.dumps({"generated_at": payload["generated_at"], "changes": changes}, ensure_ascii=False, indent=2), "utf-8")
    columns = list(Notice.__dataclass_fields__.keys())
    with CSV_PATH.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(asdict(n) for n in notices)
    HTML_PATH.parent.mkdir(parents=True, exist_ok=True)
    HTML_PATH.write_text(render_html(notices, schools, payload["generated_at"]), "utf-8")


def nl2br(text: str) -> str:
    return "<br>".join(escape(line) for line in (text or "").splitlines())


def status_class(status: str) -> str:
    return {"报名中": "open", "3天内截止": "soon", "已截止": "closed", "截止时间待核验": "review"}.get(status, "review")


def render_html(notices: list[Notice], schools: list[dict], generated_at: str) -> str:
    counts = {k: sum(n.status == k for n in notices) for k in ["报名中", "3天内截止", "已截止", "截止时间待核验"]}
    by_tier: dict[str, list[Notice]] = {"985": [], "211": []}
    for n in notices:
        by_tier.setdefault(n.tier, []).append(n)
    sections = []
    for tier in ["985", "211"]:
        items = by_tier.get(tier, [])
        school_counts = {s: sum(x.school == s for x in items) for s in {x.school for x in items}}
        seen = set()
        rows = []
        for n in items:
            first = n.school not in seen
            seen.add(n.school)
            school_cell = f'<td rowspan="{school_counts[n.school]}" class="school"><b>{escape(n.school)}</b></td>' if first else ""
            tier_cell = f'<td rowspan="{school_counts[n.school]}" class="tier">{escape(tier)}</td>' if first else ""
            group_class = " group-start" if first else ""
            rows.append(f'''<tr class="{group_class.strip()}" data-status="{escape(n.status)}" data-school="{escape(n.school)}">
            {tier_cell}{school_cell}<td><b>{escape(n.unit)}</b></td><td>{nl2br(n.programs)}</td>
            <td>{escape(n.application_start or '待核验')}<br>—<br>{escape(n.deadline or '待核验')}</td>
            <td>{nl2br(n.assessment)}</td><td><a href="{escape(n.source_url)}" target="_blank">官方公告</a></td>
            <td>{('<a href="'+escape(n.application_url)+'" target="_blank">报名系统</a>') if n.application_url else '见公告'}</td>
            <td>{nl2br(n.materials)}</td><td><span class="badge {status_class(n.status)}">{escape(n.status)}</span><br><b class="deadline">{escape(n.deadline or '待核验')}</b></td>
            <td>{escape(n.confidence)}</td></tr>''')
        sections.append(f'<h2>{tier}院校</h2><div class="wrap"><table><thead><tr><th>院校等级</th><th>院校名称</th><th>学院/项目</th><th>专业方向</th><th>报名时间</th><th>考核方式</th><th>公告链接</th><th>报名链接</th><th>申请材料清单</th><th>状态/截止</th><th>核验状态</th></tr></thead><tbody>{"".join(rows)}</tbody></table></div>')
    return f'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>2027设计类预推免监测</title>
<style>:root{{--ink:#302d32;--muted:#746e76;--line:#d9d4dc;--pink:#f5e7ed;--lav:#eee9f6;--red:#c62828}}*{{box-sizing:border-box}}body{{margin:0;background:#fff;color:var(--ink);font:14px/1.65 "Microsoft YaHei","PingFang SC",sans-serif}}main{{max-width:1900px;margin:auto;padding:24px}}h1{{margin:0}}h2{{margin:28px 0 10px;background:var(--pink);padding:8px 12px;border-left:5px solid #b85c78}}.sub{{color:var(--muted)}}.toolbar{{display:flex;gap:10px;margin:16px 0}}input,select{{padding:9px;border:1px solid var(--line)}}.wrap{{overflow:auto}}table{{width:100%;min-width:1800px;border-collapse:collapse}}th,td{{border:1px solid var(--line);padding:10px;vertical-align:top}}th{{background:var(--pink);position:sticky;top:0;z-index:2}}tr.group-start td{{border-top:3px solid #8b808d}}td.school,td.tier{{background:#fbf7fa;text-align:center;vertical-align:middle}}a{{color:#6e5b96}}.badge{{font-weight:700}}.deadline{{color:var(--red)}}.open{{color:#17643d}}.soon{{color:#b26a00}}.closed{{color:#707070}}.review{{color:#70409a}}@media(max-width:800px){{main{{padding:12px}}}}</style></head><body><main>
<h1>2027届985/211高校设计类预推免监测</h1><div class="sub">最近更新：{escape(generated_at)}｜同校连续分组｜专业、考核、材料均按条目展示｜官方站点低频轮询</div>
<div class="toolbar"><input id="q" placeholder="搜索学校、学院、专业"><select id="s"><option value="">全部状态</option><option>报名中</option><option>3天内截止</option><option>截止时间待核验</option><option>已截止</option></select></div>
{''.join(sections)}
<p class="sub">自动抽取用于发现新通知；提交申请前必须逐项打开官方公告复核。</p>
<script>const q=document.querySelector('#q'),s=document.querySelector('#s');function f(){{let x=q.value.toLowerCase(),st=s.value;document.querySelectorAll('tbody tr').forEach(r=>r.style.display=(!x||r.innerText.toLowerCase().includes(x))&&(!st||r.dataset.status===st)?'':'none')}}q.oninput=f;s.onchange=f;</script></main></body></html>'''


def notify_webhook(changes: list[dict]) -> None:
    url = os.getenv("WEBHOOK_URL", "").strip()
    if not url or not changes:
        return
    try:
        requests.post(url, json={"text": "设计类预推免监测有更新", "changes": changes[:20]}, timeout=15).raise_for_status()
    except requests.RequestException as exc:
        print(f"[WARN] webhook failed: {exc}", file=sys.stderr)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--seed-only", action="store_true", help="只用已核验种子数据生成输出，不访问网络")
    parser.add_argument("--max-pages", type=int, default=18)
    parser.add_argument("--delay", type=float, default=1.3)
    args = parser.parse_args()
    schools = json.loads(CONFIG_PATH.read_text("utf-8"))
    seed = read_seed()
    crawled: list[Notice] = []
    if not args.seed_only:
        crawler = Crawler(delay=max(0.8, args.delay), max_pages_per_source=max(3, args.max_pages))
        for school in schools:
            for url in school.get("urls", []):
                print(f"[INFO] {school['tier']} {school['school']} :: {url}")
                crawled.extend(crawler.crawl_source(school, url))
    notices, changes = merge_notices(seed, crawled)
    save_outputs(notices, changes, schools)
    notify_webhook(changes)
    print(f"[OK] notices={len(notices)} changes={len(changes)} output={HTML_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
