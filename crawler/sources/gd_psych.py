"""广东教师编制公告 · 心理岗位专项爬虫。

工作流程：
1. 抓取广东各市/区教育局、人社局的"招聘公告"列表页，挑出教师招聘公告；
2. 对每条新公告打开详情页，下载附件里的《岗位表》（xls / xlsx / doc / docx / zip）；
3. 在岗位表里查找含"心理"的行（心理健康教育、心理学专业等），记下具体岗位；
4. 结果写到 web/data/gd-psych.json，供 web/gd-psych.html 页面展示。

为什么要看附件：编制公告的标题通常只写"公开招聘教师 XX 名"，
心理老师岗位几乎都藏在附件岗位表里，只看标题会全部漏掉。

想加新网站：在下面 SITES 里照格式加一行即可（name / list_url / city，可选 max_pages 单独指定翻页数）。
"""

from __future__ import annotations

import io
import json
import os
import random
import re
import ssl
import time
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
from crawler.fetch import DEFAULT_UA, FetchError, _CTX, detect_charset, polite_delay
from crawler.sources import config as source_config
from crawler.storage import WEB_DATA_DIR, save_json

try:  # 读取老版 .xls 岗位表；没装也能运行，只是改用粗略的文本搜索
    import xlrd  # type: ignore
except ImportError:  # pragma: no cover
    xlrd = None

try:  # 模拟真实 Chrome 浏览器的网络特征（部分政府网站会拦截"不像浏览器"的访问）
    from curl_cffi import requests as _creq  # type: ignore
except ImportError:  # pragma: no cover
    _creq = None

try:  # 读取 PDF 岗位表（可选）
    from pypdf import PdfReader  # type: ignore
except ImportError:  # pragma: no cover
    PdfReader = None


# ---------------------------------------------------------------------------
# 数据源：广东各地教育局 / 人社局的招聘公告列表页
# 大部分广东政府网站用同一套系统，翻页规则是 index.html → index_2.html → index_3.html
# ---------------------------------------------------------------------------
SITES = [
    # ---- 省级 ----
    {"name": "广东省人社厅·事业单位招聘", "list_url": "https://hrss.gd.gov.cn/zwgk/sydwzp/zpgg/index.html", "city": "省直属"},
    # ---- 广州 ----
    {"name": "广州市教育局·通知公告", "list_url": "https://jyj.gz.gov.cn/yw/tzgg/index.html", "city": "广州"},
    {"name": "广州市人社局·事业单位招聘", "list_url": "https://rsj.gz.gov.cn/ywzt/rszdgg/sydwgkzp/sydwzpgg/index.html", "city": "广州"},
    {"name": "广州天河区·人事招聘", "list_url": "http://www.thnet.gov.cn/thdt/tzgg/rsxx/rszp/index.html", "city": "广州"},
    {"name": "广州天河区教育局·人事招聘", "list_url": "http://www.thnet.gov.cn/gzjg/qzf/qjyj/tzgg/rszp/index.html", "city": "广州"},
    {"name": "广州海珠区教育局·教师", "list_url": "https://www.haizhu.gov.cn/gzjg/qzf/hzqjyj/js/index.html", "city": "广州"},
    {"name": "广州海珠区·招聘信息", "list_url": "http://www.haizhu.gov.cn/hzdt/tzgg/zpxx/index.html", "city": "广州"},
    {"name": "广州番禺区·招考信息", "list_url": "https://www.panyu.gov.cn/zwgk/rsgk/zkxx/index.html", "city": "广州"},
    {"name": "广州黄埔区·招聘公告", "list_url": "http://www.hp.gov.cn/xwzx/tzgg/zpgg/index.html", "city": "广州"},
    {"name": "广州增城区教育局·通知公告", "list_url": "http://www.zc.gov.cn/jg/qzfbm/qjyj/tzgg/index.html", "city": "广州"},
    # ---- 深圳 ----
    {"name": "深圳市教育局·公办学校招聘", "list_url": "https://szeb.sz.gov.cn/home/xxgk/zthd/jszp/gbxx/index.html", "city": "深圳"},
    {"name": "深圳市教育局·人员招聘", "list_url": "https://szeb.sz.gov.cn/home/xxgk/flzy/rsxx2/ryzp/index.html", "city": "深圳"},
    {"name": "深圳福田区·招考专栏", "list_url": "https://www.szft.gov.cn/xxgk/ztbd/ftqzkzl/index.html", "city": "深圳"},
    {"name": "深圳龙华区·招考招聘", "list_url": "https://www.szlhq.gov.cn/xxgk/rsxx/zkzp/index.html", "city": "深圳"},
    {"name": "深圳龙华区教育局·招聘信息", "list_url": "http://www.szlhq.gov.cn/bmxxgk/jyj/dtxx_124232/zpxx/index.html", "city": "深圳"},
    # ---- 珠海 ----
    {"name": "珠海市教育局·人事信息", "list_url": "https://zhjy.zhuhai.gov.cn/zwgk/rsxx/index.html", "city": "珠海"},
    {"name": "珠海市教育局·教师队伍", "list_url": "https://zhjy.zhuhai.gov.cn/ywgz/jsdw/index.html", "city": "珠海"},
    {"name": "珠海市人社局·公职招考", "list_url": "https://zhrsj.zhuhai.gov.cn/zw/tzgg/gzzk/index.html", "city": "珠海"},
    {"name": "珠海市政府·公职招考", "list_url": "https://www.zhuhai.gov.cn/zw/rsxx/gzzk/index.html", "city": "珠海"},
    # ---- 东莞（这两个网站的 robots.txt 可能不允许爬取，会被自动跳过） ----
    {"name": "东莞市教育局·公示公告", "list_url": "https://edu.dg.gov.cn/jyzx/gsgg/index.html", "city": "东莞"},
    {"name": "东莞市人社局·公开招聘", "list_url": "https://dghrss.dg.gov.cn/xwzx/gsgg/gkzp/index.html", "city": "东莞"},
    # ---- 中山（教体局网站的"人事信息"只有任免通知，招聘公告都发在人社局；
    #      该栏目大量"拟聘用名单"，教师公告常排在第 10 页以后，所以多翻几页） ----
    {"name": "中山市人社局·事业单位公开招聘", "list_url": "https://hrss.zs.gov.cn/xxgk/rsxx/sydwgkzp/index.html", "city": "中山", "max_pages": 20},
    # ---- 佛山 ----
    {"name": "佛山市教育局·招聘信息", "list_url": "https://edu.foshan.gov.cn/gg/zhaopinxinxi/index.html", "city": "佛山"},
    {"name": "佛山市人社局·机关事业单位招录", "list_url": "https://hrss.foshan.gov.cn/zwgk/jgsydwzl/index.html", "city": "佛山"},
    # ---- 惠州 ----
    {"name": "惠州市教育局·人事工作", "list_url": "https://jyj.huizhou.gov.cn/zwgk/rsgz/index.html", "city": "惠州"},
    {"name": "惠州市人社局·事业单位人事管理", "list_url": "https://rsj.huizhou.gov.cn/ywzt/sydwrsgl/index.html", "city": "惠州"},
]

# 标题必须同时满足：含"招聘类"词 + 含"教师类"词，且不含"过程类"词
TITLE_RECRUIT_WORDS = ["招聘", "招考", "引进", "选聘"]
TITLE_TEACHER_WORDS = ["教师", "教职员", "教职工", "教育系统", "教体系统", "学校", "中学", "小学", "幼儿园", "心理"]
TITLE_SKIP_WORDS = [
    "拟聘", "拟录用", "名单", "成绩", "面试", "体检", "考察", "资格审查", "资格复审", "资格初审",
    "递补", "分数线", "准考证", "考场", "编外", "临聘", "劳务派遣", "结果", "聘用人员公示",
    "关于公布", "笔试", "考核公告", "考试的通知", "考试安排", "报名的通知",
]

# 岗位表里命中"心理"，但其实与心理岗位无关的常见说法，先剔除再判断
PSYCH_FALSE_POSITIVES = [
    "教育学、心理学", "教育学，心理学", "教育学,心理学", "教育学和心理学", "教育学与心理学",
    "教育学心理学", "教育学及心理学", "心理素质", "心理条件", "心理健康状况", "身体和心理", "身心健康", "运动心理学", "体育心理学", "体育与心理健康",
]

# 附件后缀与需跳过的附件（专业参考目录里必然有"心理学"，会造成误报）
ATTACH_EXTS = (".xls", ".xlsx", ".et", ".doc", ".docx", ".wps", ".zip", ".pdf")
ATTACH_SKIP_WORDS = [
    "专业参考目录", "专业目录", "操作说明", "咨询电话", "同意报考", "对照表", "资格审查", "审核资料",
    "承诺书", "报名表", "登记表", "问题的解答", "报名指南", "诚信", "授权书", "体检",
]

STATE_FILE = os.path.join(WEB_DATA_DIR, "gd-psych.json")
KEEP_DAYS = 400          # 公告保留天数
RECENT_DAYS = 90         # 只检查最近三个月发布的公告
MAX_ROWS_PER_ITEM = 30   # 每条公告最多保留多少行心理岗位
SCHEMA_VERSION = 2       # 数据格式版本：旧版记录会被重新检查一次（补上报名截止时间）
MAX_TRIES = 5            # 某公告连续打不开几天后放弃


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
            pages = max_pages if getattr(ctx, "shallow", False) else int(site.get("max_pages", max_pages))
            items = _scan_list(site, pages, delay)
            candidates.extend(items)
            site_status[site["name"]] = {"ok": True, "found": len(items), "message": ""}
            print(f"  [gd_psych] {site['name']}: 找到教师招聘公告 {len(items)} 条")
        except Exception as e:  # noqa: BLE001 - 单站失败不影响其他站
            site_status[site["name"]] = {"ok": False, "found": 0, "message": str(e)[:150]}
            print(f"  [gd_psych] {site['name']} 失败: {e}")

    # 2) 同一公告常被多个网站转载：按标题归组，最新发布的优先检查；
    #    某个网址打不开时自动换另一个转载网址
    known_by_title = {_title_key(t): a for a in known.values() for t in (a["title"], a.get("list_title")) if t}
    # 打开后发现不是招聘公告 / 超过三个月的：只记标题，不进公告列表，也不再重复打开
    skipped = {s["key"]: s for s in state.get("skipped", [])}
    recent_cutoff = _recent_cutoff()
    groups: dict[str, list[dict]] = {}
    for cand in candidates:
        groups.setdefault(_title_key(cand["title"]), []).append(cand)
    ordered = sorted(groups.items(), key=lambda kv: max(c.get("publish_date") or "" for c in kv[1]), reverse=True)

    checked_now = 0
    for key, cands in ordered:
        if key in skipped:
            continue
        list_date = max(c.get("publish_date") or "" for c in cands)
        if list_date and list_date < recent_cutoff:
            continue  # 列表页日期已超过三个月，不用打开
        old = known_by_title.get(key)
        if old and not old.get("error") and old.get("v", 1) >= SCHEMA_VERSION:
            continue
        if old and old.get("error") and old.get("tries", 0) >= MAX_TRIES:
            continue
        if checked_now >= max_new:
            break
        checked_now += 1
        tries = old.get("tries", 0) + 1 if (old and old.get("error")) else 1
        item, errors = None, []
        for cand in cands:
            trial = {**cand, "checked_at": today, "tries": tries, "v": SCHEMA_VERSION}
            try:
                _inspect_announcement(trial, delay)
                item = trial
                break
            except HostPaused:
                pass
            except Exception as e:  # noqa: BLE001
                errors.append(f"{cand['site']}: {str(e)[:120]}")
                polite_delay(delay, 0.5)
        if item is None and not errors:
            # 所有转载网址所在网站都已暂停访问：本次没真正检查，不计入失败次数，留到下次
            checked_now -= 1
            print(f"  [gd_psych] 跳过 {cands[0]['title'][:40]}（网站暂停访问，下次再查）")
            continue
        if item is None:
            item = {**cands[0], "checked_at": today, "tries": tries, "v": SCHEMA_VERSION,
                    "error": "；".join(errors)[:300], "has_psych": False, "psych_rows": []}
        if item.get("skip_reason"):
            skipped[key] = {"key": key, "title": item["title"], "url": item["url"],
                            "reason": item["skip_reason"], "checked_at": today}
            if old:
                known.pop(old["url"], None)
            print(f"  [gd_psych] 检查 {item['title'][:40]} 跳过：{item['skip_reason']}")
            polite_delay(delay, 0.5)
            continue
        if old:
            known.pop(old["url"], None)
        known[item["url"]] = item
        known_by_title[key] = item
        flag = "✅ 含心理岗" if item.get("has_psych") else ("❌ 打不开" if item.get("error") else "—")
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
        "fetch_stats": HOST_STATS,
        "announcements": announcements,
        "skipped": [s for s in skipped.values() if s.get("checked_at", today) >= cutoff],
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
            "deadline": a.get("deadline", ""),
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
            html = _fetch_text(url)
        except FetchError:
            if page == 1:
                raise
            break
        items = parse_list_html(html, url, site)
        for it in items:
            if it["url"] not in seen:
                seen.add(it["url"])
                out.append(it)
        # 本页所有文章都早于三个月就不再往后翻。注意不能"本页没有教师公告就停"：
        # 暑期前几页常被"拟聘用人员公示"占满，真正的招聘公告在第 2 页以后
        newest = _newest_date(html)
        if (newest and newest < _recent_cutoff()) or (not newest and not items):
            break  # 列表页不写日期的，沿用"本页没有教师公告就停"
        polite_delay(delay, 0.4)
    return out


def _newest_date(html: str) -> str:
    """列表页上所有文章里最新的发布日期（没有日期时返回空）。"""
    soup = BeautifulSoup(html, "html.parser")
    dates = []
    for a in soup.find_all("a", href=True):
        row = a.find_parent(["li", "tr"])
        if row:
            dates.append(parse_chinese_date(row.get_text(" ", strip=True)) or "")
    return max(dates, default="")


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
            # 市/区网站直接用网站所在城市（标题里的"武汉考点"等不是工作地点）
            "city": site["city"] if site.get("city") not in ("", "省直属") else (extract_city(title) or site.get("city", "")),
            "publish_date": date,
            "list_page": page_url,
        })
    return out


# ---------------------------------------------------------------------------
# 详情页 + 附件
# ---------------------------------------------------------------------------
def _inspect_announcement(item: dict, delay: float) -> None:
    html, final_url = _fetch_text(item["url"], referer=item.get("list_page", ""), with_url=True)
    item["url"] = final_url or item["url"]
    soup = BeautifulSoup(html, "html.parser")
    body_text = soup.get_text("\n", strip=True)

    # 列表页标题常被截断（"关于公布……（广州场）..."），用详情页的完整标题再筛一次：
    # 名单/成绩类通知的附件是应聘者个人名单，既不是岗位表，也不应展示
    meta_title = soup.find("meta", attrs={"name": re.compile("^ArticleTitle$", re.I)})
    full_title = clean_text(meta_title.get("content", "")) if meta_title else ""
    if full_title:
        if full_title != item["title"]:
            item["list_title"] = item["title"]  # 下次按列表页标题也能认出这条
            item["title"] = full_title
        if not is_teacher_recruit_title(full_title):
            item["skip_reason"] = "名单/成绩/考试安排类通知，不是招聘公告"
            return

    meta_date = soup.find("meta", attrs={"name": re.compile("PubDate", re.I)})
    if meta_date and meta_date.get("content"):
        item["publish_date"] = parse_chinese_date(meta_date["content"]) or item.get("publish_date", "")
    if not item.get("publish_date"):
        item["publish_date"] = parse_chinese_date(body_text[:3000]) or ""
    item["deadline"] = extract_reg_deadline(body_text, item.get("publish_date", "")) or ""
    if item.get("publish_date") and item["publish_date"] < _recent_cutoff():
        item["skip_reason"] = "发布超过三个月"  # 列表页没写日期的，打开后才知道，不再下载附件
        return

    rows: list[str] = []
    checked: list[str] = []
    failed: list[str] = []
    for name, link in find_attachments(soup, item["url"]):
        if not _robots_allowed(link):
            continue
        try:
            data = _fetch_bytes(link, referer=item["url"])
        except HostPaused:
            raise
        except FetchError:
            checked.append(f"{name}（下载失败）")
            failed.append(name)
            continue
        checked.append(name)
        rows.extend(psych_rows_from_file(name, data))
        polite_delay(delay, 0.3)

    # 有些公告直接把岗位表放在网页正文的表格里
    for tr in soup.find_all("tr"):
        line = _join(td.get_text(" ", strip=True) for td in tr.find_all(["td", "th"]))
        if psych_hit(line):
            rows.append(line)
    # 没有附件时，逐行看正文（如"招聘心理教师 2 名"）
    if not checked:
        for line in body_text.split("\n"):
            line = clean_text(line)
            if 4 <= len(line) <= 240 and psych_hit(line):
                rows.append(line)

    if psych_hit(item["title"]) and not rows:
        rows.append(item["title"])
    # 岗位表没下载成功又没找到心理岗：可能是漏看，记为失败，下次再试（否则会被当成"已检查、无心理岗"）
    if failed and not rows:
        raise FetchError(f"附件下载失败: {'、'.join(failed)[:80]}")

    item["attachments_checked"] = checked[:10]
    item["psych_rows"] = _dedupe(rows)[:MAX_ROWS_PER_ITEM]
    item["has_psych"] = bool(item["psych_rows"])
    item.pop("note", None)
    if not checked and not item["has_psych"]:
        item["note"] = "没找到可读取的附件，建议打开原文人工查看"


def _recent_cutoff() -> str:
    return (datetime.now(timezone.utc) - timedelta(days=RECENT_DAYS)).strftime("%Y-%m-%d")


_DATE = r"(?:(20\d{2})\s*年)?\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日"
_REG_RANGE = re.compile(r"报名(?:时间|日期|期限)?[^。；;\n]{0,20}?" + _DATE + r"[^。；;\n]{0,30}?(?:至|到|—|－|-|~|～)\s*" + _DATE)
_REG_END = re.compile(r"报名截止(?:时间|日期)?[^。；;\n]{0,10}?" + _DATE)


def extract_reg_deadline(text: str, publish_date: str = "") -> str | None:
    """从正文里找报名截止日期，如"报名时间：2026年9月15日9:00至9月22日17:00" → 2026-09-22。"""
    pub_year = int(publish_date[:4]) if publish_date[:4].isdigit() else datetime.now(timezone.utc).year
    m = _REG_RANGE.search(text)
    if m:
        y1, m1, _d1, y2, m2, d2 = m.groups()
        year = int(y2 or y1 or pub_year)
        if not y2 and not y1 and publish_date[5:7].isdigit() and int(m2) < int(publish_date[5:7]):
            year += 1  # 12 月发布、次年 1 月截止
        return _iso(year, m2, d2)
    m = _REG_END.search(text)
    if m:
        y, mo, d = m.groups()
        return _iso(int(y or pub_year), mo, d)
    return None


def _iso(y, m, d) -> str | None:
    try:
        return datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d")
    except ValueError:
        return None


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


def _legacy_ssl_ctx() -> ssl.SSLContext:
    """兼容老旧政府网站的 TLS 设置（解决 BAD_ECPOINT 等握手错误）。"""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    for fn in (lambda: ctx.set_ciphers("DEFAULT:@SECLEVEL=1"),
               lambda: setattr(ctx, "maximum_version", ssl.TLSVersion.TLSv1_2),
               lambda: ctx.set_ecdh_curve("prime256v1")):
        try:
            fn()
        except (ValueError, ssl.SSLError, AttributeError):
            pass
    return ctx


_LEGACY_CTX = _legacy_ssl_ctx()


def _swap_scheme(url: str) -> str:
    if url.startswith("https://"):
        return "http://" + url[len("https://"):]
    if url.startswith("http://"):
        return "https://" + url[len("http://"):]
    return url


_CSESSION = None
HOST_STATS: dict[str, dict] = {}   # 每个网站用哪种方式成功/失败，保存到结果文件里方便排查


def _session():
    """浏览器模拟会话：会保存 Cookie，访问详情页时像真人一样"先看列表再点进去"。"""
    global _CSESSION
    if _CSESSION is None and _creq is not None:
        try:
            _CSESSION = _creq.Session(impersonate="chrome")
        except Exception:  # noqa: BLE001
            _CSESSION = False
    return _CSESSION or None


def _stat(url: str, key: str, err: Exception | None = None) -> None:
    host = urlparse(url).netloc
    st = HOST_STATS.setdefault(host, {"browser_ok": 0, "urllib_ok": 0, "fail": 0, "last_error": ""})
    st[key] += 1
    if err is not None:
        st["last_error"] = str(err)[:160]


# 个别网站对访问频率很敏感，单独放慢、少重试。
# 珠海：短时间内请求多了（尤其是失败后连续重试）会封 IP 数小时，三个子站共用一套防火墙，按整个域名合并计算。
HOST_POLICIES = {
    "zhuhai.gov.cn": {"min_interval": 8.0, "max_attempts": 2, "max_fails": 3},
}
# min_interval：同一网站两次请求的最短间隔（秒）；max_attempts：每个网址最多换几种方式尝试；
# max_fails：本次运行里连续失败几个网址后暂停访问该网站（None 表示不暂停）
DEFAULT_POLICY = {"min_interval": 1.5, "max_attempts": 6, "max_fails": None}


class HostPaused(FetchError):
    """该网站本次运行里连续失败太多，已暂停访问（没有真正发出请求）。"""


_HOST_STATE: dict[str, dict] = {}


def _host_policy(url: str) -> tuple[str, dict]:
    host = urlparse(url).netloc.split(":")[0].lower()
    for suffix, policy in HOST_POLICIES.items():
        if host == suffix or host.endswith("." + suffix):
            return suffix, {**DEFAULT_POLICY, **policy}
    return host, DEFAULT_POLICY


def _throttle(key: str, min_interval: float) -> None:
    st = _HOST_STATE.setdefault(key, {"last": 0.0, "fails": 0})
    wait = st["last"] + min_interval - time.time()
    if wait > 0:
        time.sleep(wait + random.uniform(0, 0.5))
    st["last"] = time.time()


def _request(url: str, referer: str = "", max_bytes: int = 20_000_000, timeout: int = 30):
    """依次尝试：浏览器模拟 → 普通方式 → 换 http/https → 兼容模式 TLS。返回 (内容, 最终网址, 编码)。

    每个网站的请求间隔、尝试次数、连续失败后是否暂停，见 HOST_POLICIES。
    """
    key, policy = _host_policy(url)
    st = _HOST_STATE.setdefault(key, {"last": 0.0, "fails": 0})
    if policy["max_fails"] and st["fails"] >= policy["max_fails"]:
        raise HostPaused(f"{key} 本次已连续失败 {st['fails']} 次，暂停访问: {url}")

    alt = _swap_scheme(url)
    sess = _session()
    attempts: list[tuple[str, str, ssl.SSLContext | None]] = []
    for u, ctx in ((url, _CTX), (alt, _CTX), (url, _LEGACY_CTX), (alt, _LEGACY_CTX)):
        if sess is not None and ctx is _CTX:
            attempts.append(("browser", u, None))
        if ctx is _CTX or u.startswith("https://"):  # http 不涉及 TLS，兼容模式不必重复试
            attempts.append(("urllib", u, ctx))

    last: Exception | None = None
    gone = False
    for kind, u, ctx in attempts[: policy["max_attempts"]]:
        _throttle(key, policy["min_interval"])
        headers = {"Accept-Language": "zh-CN,zh;q=0.9"}
        if referer:
            headers["Referer"] = referer
        try:
            if kind == "browser":
                r = sess.get(u, headers=headers, timeout=timeout, verify=False, allow_redirects=True)
                if r.status_code == 200:
                    _stat(url, "browser_ok")
                    st["fails"] = 0
                    m = re.search(r"charset=([\w-]+)", r.headers.get("content-type", ""), re.I)
                    return r.content[:max_bytes], str(r.url), (m.group(1) if m else None)
                last = Exception(f"HTTP {r.status_code}")
                gone = r.status_code in (404, 410)
            else:
                headers["User-Agent"] = DEFAULT_UA
                headers["Accept"] = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
                req = urllib.request.Request(u, headers=headers)
                with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
                    _stat(url, "urllib_ok")
                    st["fails"] = 0
                    return resp.read(max_bytes), resp.geturl(), resp.headers.get_content_charset()
        except urllib.error.HTTPError as e:
            last = e
            gone = e.code in (404, 410)
        except Exception as e:  # noqa: BLE001 - 连接被重置 / TLS 握手失败等，换方式再试
            last = e
        if gone:
            break
    if not gone:  # 404 是网址本身失效，不算网站拦截
        st["fails"] += 1
    _stat(url, "fail", last)
    raise FetchError(f"GET 失败: {url} -> {last}")


def _fetch_text(url: str, referer: str = "", with_url: bool = False):
    raw, final_url, enc = _request(url, referer=referer, max_bytes=5_000_000)
    enc = enc or detect_charset(raw) or "utf-8"
    try:
        text = raw.decode(enc, "ignore")
    except LookupError:
        text = raw.decode("utf-8", "ignore")
    return (text, final_url) if with_url else text


def _fetch_bytes(url: str, referer: str = "") -> bytes:
    return _request(url, referer=referer)[0]


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


def _title_key(title: str) -> str:
    return re.sub(r"[\s（）()“”\"'《》【】\[\]·、，,:：-]", "", title)


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
        key, policy = _host_policy(url)
        _throttle(key, policy["min_interval"])
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
