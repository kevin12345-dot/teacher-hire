"""广东教师编制公告 · 心理岗位专项爬虫。

工作流程：
1. 抓取广东各市/区教育局、人社局的"招聘公告"列表页，挑出教师招聘公告；
2. 对每条新公告打开详情页，下载附件里的《岗位表》（xls / xlsx / doc / docx / zip）；
3. 在岗位表里查找含"心理"的行（心理健康教育、心理学专业等），记下具体岗位；
4. 结果写到 web/data/gd-psych.json，供 web/gd-psych.html 页面展示。

为什么要看附件：编制公告的标题通常只写"公开招聘教师 XX 名"，
心理老师岗位几乎都藏在附件岗位表里，只看标题会全部漏掉。

想加新网站：在下面 SITES 里照格式加一行即可（name / list_url / city）。
"""

from __future__ import annotations

import io
import json
import os
import re
import urllib.error
import urllib.request
import zipfile
from datetime import datetime, timedelta, timezone
from hashlib import md5
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser
from xml.etree import ElementTree as ET

from bs4 import BeautifulSoup

from crawler.clean import clean_text, extract_city, parse_chinese_date
from crawler.fetch import DEFAULT_UA, FetchError, _CTX, http_get, polite_delay
from crawler.sources import config as source_config
from crawler.storage import WEB_DATA_DIR, save_json

try:  # 读取老版 .xls 岗位表；没装也能运行，只是改用粗略的文本搜索
    import xlrd  # type: ignore
except ImportError:  # pragma: no cover
    xlrd = None

try:  # 读取 PDF 岗位表（可选）
    from pypdf import PdfReader  # type: ignore
except ImportError:  # pragma: no cover
    PdfReader = None


# ---------------------------------------------------------------------------
# 数据源：广东各地教育局 / 人社局的招聘公告列表页
# 大部分广东政府网站用同一套系统，翻页规则是 index.html → index_2.html → index_3.html
# ---------------------------------------------------------------------------
SITES = [
    {"name": "广州市教育局·通知公告", "list_url": "https://jyj.gz.gov.cn/yw/tzgg/index.html", "city": "广州"},
    {"name": "广州番禺区·招考信息", "list_url": "https://www.panyu.gov.cn/zwgk/rsgk/zkxx/index.html", "city": "广州"},
    {"name": "广州黄埔区·招聘公告", "list_url": "http://www.hp.gov.cn/xwzx/tzgg/zpgg/index.html", "city": "广州"},
    {"name": "广州增城区教育局·通知公告", "list_url": "http://www.zc.gov.cn/jg/qzfbm/qjyj/tzgg/index.html", "city": "广州"},
    {"name": "深圳市教育局·公办学校招聘", "list_url": "https://szeb.sz.gov.cn/home/xxgk/zthd/jszp/gbxx/index.html", "city": "深圳"},
    {"name": "东莞市教育局·公示公告", "list_url": "https://edu.dg.gov.cn/jyzx/gsgg/index.html", "city": "东莞"},
    {"name": "东莞市人社局·公开招聘", "list_url": "https://dghrss.dg.gov.cn/xwzx/gsgg/gkzp/index.html", "city": "东莞"},
]

# 标题必须同时满足：含"招聘类"词 + 含"教师类"词，且不含"过程类"词
TITLE_RECRUIT_WORDS = ["招聘", "招考", "引进", "选聘"]
TITLE_TEACHER_WORDS = ["教师", "教职员", "教职工", "教育系统", "学校", "中学", "小学", "幼儿园", "心理"]
TITLE_SKIP_WORDS = [
    "拟聘", "拟录用", "名单", "成绩", "面试", "体检", "考察", "资格审查", "资格复审", "资格初审",
    "递补", "分数线", "准考证", "考场", "编外", "临聘", "劳务派遣", "结果", "聘用人员公示",
]

# 岗位表里命中"心理"，但其实与心理岗位无关的常见说法，先剔除再判断
PSYCH_FALSE_POSITIVES = [
    "教育学、心理学", "教育学，心理学", "教育学,心理学", "教育学和心理学", "教育学与心理学",
    "教育学心理学", "教育学及心理学", "心理素质", "心理条件", "心理健康状况", "身体和心理", "身心健康",
]

# 附件后缀与需跳过的附件（专业参考目录里必然有"心理学"，会造成误报）
ATTACH_EXTS = (".xls", ".xlsx", ".et", ".doc", ".docx", ".wps", ".zip", ".pdf")
ATTACH_SKIP_WORDS = [
    "专业参考目录", "专业目录", "操作说明", "咨询电话", "同意报考", "对照表", "资格审查", "审核资料",
    "承诺书", "报名表", "登记表", "问题的解答", "报名指南", "诚信", "授权书", "体检",
]

STATE_FILE = os.path.join(WEB_DATA_DIR, "gd-psych.json")
KEEP_DAYS = 400          # 公告保留天数
MAX_ROWS_PER_ITEM = 30   # 每条公告最多保留多少行心理岗位


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------
def run(ctx) -> list[dict]:
    cfg = source_config.SOURCES.get("gd_psych", {})
    delay = float(cfg.get("delay", 1.5))
    max_pages = 1 if getattr(ctx, "shallow", False) else int(cfg.get("max_pages", 3))
    max_new = int(cfg.get("max_new_per_run", 40))

    state = _load_state()
    known = {a["url"]: a for a in state.get("announcements", [])}
    site_status: dict[str, dict] = {}
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    # 1) 扫列表页，收集候选公告
    candidates: list[dict] = []
    for site in SITES:
        try:
            if not _robots_allowed(site["list_url"]):
                site_status[site["name"]] = {"ok": False, "found": 0, "message": "该网站 robots.txt 不允许爬取，已跳过"}
                print(f"  [gd_psych] {site['name']}: robots 不允许，跳过")
                continue
            items = _scan_list(site, max_pages, delay)
            candidates.extend(items)
            site_status[site["name"]] = {"ok": True, "found": len(items), "message": ""}
            print(f"  [gd_psych] {site['name']}: 找到教师招聘公告 {len(items)} 条")
        except Exception as e:  # noqa: BLE001 - 单站失败不影响其他站
            site_status[site["name"]] = {"ok": False, "found": 0, "message": str(e)[:150]}
            print(f"  [gd_psych] {site['name']} 失败: {e}")

    # 2) 对没检查过（或上次失败）的公告，打开详情页和附件检查
    checked_now = 0
    for cand in candidates:
        old = known.get(cand["url"])
        if old and not old.get("error"):
            continue
        if old and old.get("tries", 0) >= 3:
            continue
        if checked_now >= max_new:
            break
        checked_now += 1
        item = {**cand, "checked_at": today, "tries": (old or {}).get("tries", 0) + 1}
        try:
            _inspect_announcement(item, delay)
            item.pop("error", None)
        except Exception as e:  # noqa: BLE001
            item["error"] = str(e)[:150]
            item.setdefault("has_psych", False)
            item.setdefault("psych_rows", [])
        known[cand["url"]] = item
        flag = "✅ 含心理岗" if item.get("has_psych") else "—"
        print(f"  [gd_psych] 检查 {item['title'][:40]} {flag}")
        polite_delay(delay, 0.5)

    # 3) 清理过旧公告并保存
    cutoff = (datetime.now(timezone.utc) - timedelta(days=KEEP_DAYS)).strftime("%Y-%m-%d")
    announcements = [a for a in known.values() if (a.get("publish_date") or a.get("checked_at") or today) >= cutoff]
    announcements.sort(key=lambda a: (a.get("publish_date") or "", a.get("checked_at") or ""), reverse=True)
    state = {
        "updated_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "latest_check_date": today,
        "sites": site_status,
        "announcements": announcements,
    }
    save_json(STATE_FILE, state)

    # 4) 把含心理岗位的公告交给主站（统一字段）
    records = []
    for a in announcements:
        if not a.get("has_psych"):
            continue
        rows = a.get("psych_rows", [])
        summary = f"含心理相关岗位 {len(rows)} 个：" + ("；".join(rows[:2]))[:130]
        records.append({
            "id": md5(a["url"].encode("utf-8")).hexdigest()[:16],
            "title": a["title"],
            "school": a.get("site", ""),
            "url": a["url"],
            "source": "gd_psych",
            "source_label": source_config.SOURCE_LABELS.get("gd_psych", "广东官网·心理岗"),
            "province": "广东",
            "city": a.get("city", ""),
            "education": "",
            "experience": "",
            "salary": "",
            "salary_text": "事业编制",
            "subject": "心理",
            "school_level": "",
            "deadline": "",
            "publish_date": a.get("publish_date", ""),
            "crawl_date": "",
            "summary": summary,
        })
    print(f"  [gd_psych] 本次新检查 {checked_now} 条，累计含心理岗公告 {len(records)} 条")
    return records


# ---------------------------------------------------------------------------
# 列表页
# ---------------------------------------------------------------------------
def _scan_list(site: dict, max_pages: int, delay: float) -> list[dict]:
    out: list[dict] = []
    seen: set[str] = set()
    for page in range(1, max_pages + 1):
        url = _page_url(site["list_url"], page)
        try:
            html = http_get(url, delay=delay, timeout=25)
        except FetchError:
            if page == 1:
                raise
            break
        items = parse_list_html(html, url, site)
        if not items:
            break
        for it in items:
            if it["url"] not in seen:
                seen.add(it["url"])
                out.append(it)
        polite_delay(delay, 0.4)
    return out


def _page_url(list_url: str, page: int) -> str:
    if page == 1:
        return list_url
    if list_url.endswith("index.html"):
        return list_url[: -len("index.html")] + f"index_{page}.html"
    if list_url.endswith("/"):
        return f"{list_url}index_{page}.html"
    return f"{list_url}?page={page}"


def is_teacher_recruit_title(title: str) -> bool:
    if not any(w in title for w in TITLE_RECRUIT_WORDS):
        return False
    if not any(w in title for w in TITLE_TEACHER_WORDS):
        return False
    return not any(w in title for w in TITLE_SKIP_WORDS)


def parse_list_html(html: str, page_url: str, site: dict) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith(("javascript", "#", "mailto")):
            continue
        title = clean_text(a.get("title") or a.get_text(" ", strip=True))
        if len(title) < 8 or len(title) > 120:
            continue
        if not is_teacher_recruit_title(title):
            continue
        url = urljoin(page_url, href)
        path = urlparse(url).path
        if not ("/content/" in path or "post_" in path or re.search(r"/t20\d{6}_", path)
                or path.endswith((".html", ".shtml", ".htm"))):
            continue
        # 日期一般写在同一行（li / tr）里
        row_text = a.find_parent(["li", "tr", "div"])
        date = parse_chinese_date(row_text.get_text(" ", strip=True) if row_text else "") or ""
        out.append({
            "id": md5(url.encode("utf-8")).hexdigest()[:16],
            "title": title,
            "url": url,
            "site": site["name"],
            "city": extract_city(title) or site.get("city", ""),
            "publish_date": date,
        })
    return out


# ---------------------------------------------------------------------------
# 详情页 + 附件
# ---------------------------------------------------------------------------
def _inspect_announcement(item: dict, delay: float) -> None:
    html = http_get(item["url"], delay=delay, timeout=25)
    soup = BeautifulSoup(html, "html.parser")

    meta_date = soup.find("meta", attrs={"name": re.compile("PubDate", re.I)})
    if meta_date and meta_date.get("content"):
        item["publish_date"] = parse_chinese_date(meta_date["content"]) or item.get("publish_date", "")
    if not item.get("publish_date"):
        item["publish_date"] = parse_chinese_date(soup.get_text(" ", strip=True)[:3000]) or ""

    rows: list[str] = []
    checked: list[str] = []
    for name, link in find_attachments(soup, item["url"]):
        if not _robots_allowed(link):
            continue
        try:
            data = http_get_bytes(link)
        except FetchError:
            checked.append(f"{name}（下载失败）")
            continue
        checked.append(name)
        rows.extend(psych_rows_from_file(name, data))
        polite_delay(delay, 0.3)

    # 标题本身就写了心理岗（如"招聘心理健康教育教师"）
    if psych_hit(item["title"]) and not rows:
        rows.append(item["title"])

    item["attachments_checked"] = checked[:10]
    item["psych_rows"] = _dedupe(rows)[:MAX_ROWS_PER_ITEM]
    item["has_psych"] = bool(item["psych_rows"])
    if not checked:
        item["note"] = "没找到可读取的附件，建议打开原文人工查看"


def find_attachments(soup: BeautifulSoup, page_url: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        link = urljoin(page_url, a["href"].strip())
        path = urlparse(link).path.lower()
        text = clean_text(a.get("title") or a.get_text(" ", strip=True))
        name = text or os.path.basename(path)
        if not (path.endswith(ATTACH_EXTS) or name.lower().endswith(ATTACH_EXTS)):
            continue
        if any(w in name for w in ATTACH_SKIP_WORDS):
            continue
        if link in seen:
            continue
        seen.add(link)
        if not name.lower().endswith(ATTACH_EXTS):
            name += os.path.splitext(path)[1]
        found.append((name, link))
    return found[:8]


def http_get_bytes(url: str, timeout: int = 40, max_bytes: int = 20_000_000, retries: int = 2) -> bytes:
    last = None
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": DEFAULT_UA, "Accept-Language": "zh-CN,zh;q=0.9"})
            with urllib.request.urlopen(req, timeout=timeout, context=_CTX) as resp:
                return resp.read(max_bytes)
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
            last = e
            polite_delay(1.5 * (attempt + 1), 0.5)
    raise FetchError(f"下载失败: {url} -> {last}")


# ---------------------------------------------------------------------------
# 从岗位表中找"心理"行
# ---------------------------------------------------------------------------
def psych_hit(text: str) -> bool:
    if not text or "心理" not in text:
        return False
    t = text
    for fp in PSYCH_FALSE_POSITIVES:
        t = t.replace(fp, "")
    return "心理" in t


def psych_rows_from_file(name: str, data: bytes, depth: int = 0) -> list[str]:
    ext = os.path.splitext(name.lower())[1]
    try:
        if data[:2] == b"PK":  # zip 格式：可能是压缩包，也可能是改了后缀的 xlsx/docx
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = set(zf.namelist())
            if "xl/workbook.xml" in names:
                ext = ".xlsx"
            elif "word/document.xml" in names:
                ext = ".docx"
            else:
                return _rows_from_zip(data, depth)
        if ext == ".xlsx":
            return [r for r in _xlsx_rows(data) if psych_hit(r)]
        if ext == ".docx":
            return [r for r in _docx_rows(data) if psych_hit(r)]
        if ext in (".xls", ".et") and xlrd is not None:
            try:
                return [r for r in _xls_rows(data) if psych_hit(r)]
            except Exception:  # noqa: BLE001 - 有些 .xls 其实是 html/xlsx，退回文本搜索
                pass
        if ext == ".pdf":
            return [r for r in _pdf_lines(data) if psych_hit(r)]
        return _binary_snippets(data)
    except Exception as e:  # noqa: BLE001
        print(f"    读取附件 {name} 出错: {e}")
        return []


def _rows_from_zip(data: bytes, depth: int) -> list[str]:
    if depth > 1:
        return []
    rows: list[str] = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            inner = info.filename
            try:  # 中文 Windows 打的压缩包，文件名通常是 GBK 编码
                inner = inner.encode("cp437").decode("gbk")
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
            base = os.path.basename(inner)
            if not base.lower().endswith(ATTACH_EXTS) or any(w in base for w in ATTACH_SKIP_WORDS):
                continue
            rows.extend(psych_rows_from_file(base, zf.read(info), depth + 1))
    return rows


_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _join(cells) -> str:
    parts = [clean_text(str(int(c)) if isinstance(c, float) and c.is_integer() else str(c)) for c in cells]
    parts = [p for p in parts if p and p.lower() != "nan"]
    return " ｜ ".join(parts)[:240]


def _xlsx_rows(data: bytes) -> list[str]:
    rows: list[str] = []
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.iter(f"{_NS_MAIN}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{_NS_MAIN}t")))
        sheets = sorted(n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
        for sheet in sheets:
            root = ET.fromstring(zf.read(sheet))
            for row in root.iter(f"{_NS_MAIN}row"):
                cells = []
                for c in row.iter(f"{_NS_MAIN}c"):
                    t = c.get("t")
                    if t == "inlineStr":
                        cells.append("".join(x.text or "" for x in c.iter(f"{_NS_MAIN}t")))
                        continue
                    v = c.find(f"{_NS_MAIN}v")
                    if v is None or v.text is None:
                        continue
                    if t == "s":
                        idx = int(v.text)
                        cells.append(shared[idx] if idx < len(shared) else "")
                    else:
                        cells.append(v.text)
                line = _join(cells)
                if line:
                    rows.append(line)
    return rows


def _xls_rows(data: bytes) -> list[str]:
    book = xlrd.open_workbook(file_contents=data)
    rows = []
    for sheet in book.sheets():
        for i in range(sheet.nrows):
            line = _join(sheet.row_values(i))
            if line:
                rows.append(line)
    return rows


def _docx_rows(data: bytes) -> list[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        root = ET.fromstring(zf.read("word/document.xml"))
    body = root.find(f"{_NS_W}body")
    rows = []
    for child in (body if body is not None else []):
        if child.tag == f"{_NS_W}tbl":
            for tr in child.iter(f"{_NS_W}tr"):
                cells = ["".join(t.text or "" for t in tc.iter(f"{_NS_W}t")) for tc in tr.iter(f"{_NS_W}tc")]
                line = _join(cells)
                if line:
                    rows.append(line)
        elif child.tag == f"{_NS_W}p":
            line = clean_text("".join(t.text or "" for t in child.iter(f"{_NS_W}t")))
            if line:
                rows.append(line[:240])
    return rows


def _pdf_lines(data: bytes) -> list[str]:
    if PdfReader is None:
        return []
    reader = PdfReader(io.BytesIO(data))
    lines = []
    for page in reader.pages[:40]:
        for line in (page.extract_text() or "").splitlines():
            line = clean_text(line)
            if line:
                lines.append(line[:240])
    return lines


def _binary_snippets(data: bytes) -> list[str]:
    """老版 .doc / .xls / .wps：文字一般以 UTF-16 存储，直接在字节里找"心理"前后的文字。"""
    out = []
    for enc in ("utf-16-le", "gb18030", "utf-8"):
        text = data.decode(enc, "ignore")
        text = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9（）()、，,：:；;/\-\s]", " ", text)
        for m in re.finditer("心理", text):
            snippet = clean_text(text[max(0, m.start() - 30): m.end() + 30])
            if psych_hit(snippet):
                out.append(snippet)
        if out:
            break
    return out


def _dedupe(rows: list[str]) -> list[str]:
    seen, out = set(), []
    for r in rows:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out


# ---------------------------------------------------------------------------
# robots.txt 与状态文件
# ---------------------------------------------------------------------------
_ROBOTS: dict[str, RobotFileParser | None] = {}


def _robots_allowed(url: str) -> bool:
    parts = urlparse(url)
    base = f"{parts.scheme}://{parts.netloc}"
    if base not in _ROBOTS:
        rp = RobotFileParser()
        try:
            req = urllib.request.Request(base + "/robots.txt", headers={"User-Agent": DEFAULT_UA})
            with urllib.request.urlopen(req, timeout=10, context=_CTX) as resp:
                rp.parse(resp.read(200_000).decode("utf-8", "ignore").splitlines())
            _ROBOTS[base] = rp
        except urllib.error.HTTPError as e:
            # 401/403：按惯例视为禁止；404 等：视为没有限制
            _ROBOTS[base] = False if e.code in (401, 403) else None
        except Exception:  # noqa: BLE001 - 取不到 robots.txt 时不阻塞
            _ROBOTS[base] = None
    rp = _ROBOTS[base]
    if rp is False:
        return False
    if rp is None:
        return True
    return rp.can_fetch("*", url)


def _load_state() -> dict:
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}
