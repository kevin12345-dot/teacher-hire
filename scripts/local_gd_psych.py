"""本机每日运行 gd_psych（广东心理岗），并把 web/data/gd-psych.json 提交推送到 GitHub。

为什么要在本机跑：珠海、东莞等网站会拦截 GitHub 云服务器（境外 IP），国内网络可以正常访问。

流程：
1. git pull --rebase 拉取最新代码和数据（云端每天 15:00 也会更新同一个文件）；
2. 检查政府网站能否直连（代理把 gov.cn 送出境外时会全部失败，此时直接停止）；
3. 只运行 gd_psych；
4. 只提交 web/data/gd-psych.json 并推送；推送被拒（云端刚好也推了）时，
   拉取远端版本，按公告逐条合并后再推。

用法（一般由 scripts/daily_local_run.bat 调用）：
    python scripts/local_gd_psych.py              # 正常运行：抓取 + 提交 + 推送
    python scripts/local_gd_psych.py --no-push    # 抓取 + 本地提交，不推送
    python scripts/local_gd_psych.py --dry-run    # 只抓取，结果写到 logs/ 下的副本，不动仓库
    python scripts/local_gd_psych.py --skip-city 珠海   # 本次不访问某城市（可重复）
    python scripts/local_gd_psych.py --only-if-due      # 今天已成功跑过就跳过（计划任务用）

计划任务：每天 15:00 触发，另在"登录后 30 分钟""睡眠唤醒后 30 分钟"触发（错过 15:00 时补跑），
三个触发都带 --only-if-due：每天只由最先到的那个运行一次，其余跳过。
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from crawler.sources import gd_psych  # noqa: E402
from crawler.storage import save_json  # noqa: E402

DATA_FILE = "web/data/gd-psych.json"
LOG_DIR = os.path.join(ROOT, "logs")
# 用来判断"政府网站能否直连"：广东省人社厅在境外 IP 下打不开，国内直连正常
PROBE_URL = "https://hrss.gd.gov.cn/zwgk/sydwzp/zpgg/index.html"
PUSH_RETRIES = 3   # 推送被拒（云端同时更新）时合并重推的次数
NET_RETRIES = 5    # 连不上 GitHub 时的重试次数
SUCCESS_FILE = os.path.join(LOG_DIR, "last_success.txt")  # 最近一次完整成功运行的时间

_log_file = None


def log(msg: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line, flush=True)
    if _log_file:
        _log_file.write(line + "\n")
        _log_file.flush()


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess:
    r = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if check and r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 失败：{(r.stderr or r.stdout).strip()[:500]}")
    return r


# 连不上 GitHub 时的报错特征（开机时代理还没启动、网络抖动等）
_NET_ERRORS = ("Could not connect", "Failed to connect", "timed out", "RPC failed", "Could not resolve",
               "Connection was reset", "unable to access", "early EOF")


def git_net(*args: str, check: bool = True) -> subprocess.CompletedProcess:
    """需要联网的 git 命令（pull / push）：连不上 GitHub 时每隔 1 分钟重试，最多重试 NET_RETRIES 次。"""
    for i in range(NET_RETRIES + 1):
        r = git(*args, check=False)
        err = r.stderr or ""
        if r.returncode == 0 or not any(s in err for s in _NET_ERRORS) or i == NET_RETRIES:
            break
        log(f"连不上 GitHub（git {args[0]}），1 分钟后重试（{i + 1}/{NET_RETRIES}）：{err.strip()[:120]}")
        time.sleep(60)
    if check and r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} 失败：{(r.stderr or r.stdout).strip()[:500]}")
    return r


def is_due(now: datetime) -> bool:
    """今天还没成功运行过才需要运行：15:00、登录、唤醒三个触发里，每天只由最先到的那个运行一次。"""
    try:
        with open(SUCCESS_FILE, encoding="utf-8") as f:
            last = datetime.fromisoformat(f.read().strip())
    except (OSError, ValueError):
        return True
    return last.date() < now.date()


def mark_success() -> None:
    with open(SUCCESS_FILE, "w", encoding="utf-8") as f:
        f.write(datetime.now().isoformat(timespec="seconds"))


def current_branch() -> str:
    return git("rev-parse", "--abbrev-ref", "HEAD").stdout.strip()


# ---------------------------------------------------------------------------
# 网络检查
# ---------------------------------------------------------------------------
def check_gov_reachable() -> None:
    try:
        # 走 gd_psych 同一套请求方式（浏览器模拟 → 普通方式），结果与正式抓取一致
        gd_psych._fetch_text(PROBE_URL)
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(
            f"政府网站打不开（{str(e)[:150]}）。请检查 FlClash 是否开启了 DOMAIN-SUFFIX,gov.cn,DIRECT 规则，"
            "或者当前网络能否访问 hrss.gd.gov.cn。本次不运行，避免把失败次数记到数据里。"
        ) from e


# ---------------------------------------------------------------------------
# 两份 gd-psych.json 合并（本机和云端同时更新时用）
# ---------------------------------------------------------------------------
def _rank(a: dict) -> tuple:
    # 成功检查过的优先于"打不开"的；同类里检查日期新的优先
    return (0 if a.get("error") else 1, a.get("checked_at") or "", a.get("v", 1))


def merge_states(local: dict, remote: dict) -> dict:
    key = gd_psych._title_key
    merged: dict[str, dict] = {}
    for a in remote.get("announcements", []) + local.get("announcements", []):  # 本机在后：同等情况下本机优先
        k = key(a.get("list_title") or a["title"])
        old = merged.get(k)
        if old is None or _rank(a) >= _rank(old):
            if old and old.get("error") and a.get("error"):
                a = {**a, "tries": max(a.get("tries", 0), old.get("tries", 0))}
            merged[k] = a
    skipped = {s["key"]: s for s in remote.get("skipped", []) + local.get("skipped", [])}
    announcements = [a for k, a in merged.items()
                     if not (k in skipped or key(a["title"]) in skipped) or a.get("has_psych")]
    announcements.sort(key=lambda a: (a.get("publish_date") or "", a.get("checked_at") or ""), reverse=True)
    out = {**remote, **local}
    out["announcements"] = announcements
    out["skipped"] = list(skipped.values())
    return out


# ---------------------------------------------------------------------------
# 提交与推送
# ---------------------------------------------------------------------------
def commit_data(message: str) -> bool:
    git("add", DATA_FILE)
    if git("diff", "--cached", "--quiet", "--", DATA_FILE, check=False).returncode == 0:
        return False
    git("commit", "-m", message, "--", DATA_FILE)
    return True


def push_with_merge(message: str) -> None:
    branch = current_branch()
    for attempt in range(1, PUSH_RETRIES + 1):
        r = git_net("push", "origin", branch, check=False)
        if r.returncode != 0 and any(s in (r.stderr or "") for s in _NET_ERRORS):
            raise RuntimeError(f"连不上 GitHub，数据已提交在本地，下次运行会一起推送：{r.stderr.strip()[:200]}")
        if r.returncode == 0:
            log(f"推送成功（{git('rev-parse', '--short', 'HEAD').stdout.strip()}）")
            return
        log(f"推送被拒（第 {attempt} 次），可能云端刚好也更新了数据，拉取后合并再试：{r.stderr.strip()[:200]}")
        with open(os.path.join(ROOT, DATA_FILE), encoding="utf-8") as f:
            local_state = json.load(f)
        # 撤销本次数据提交，拿远端最新版本，再把本机结果合并进去
        git("reset", "--soft", "HEAD~1")
        git("restore", "--staged", DATA_FILE)
        git("checkout", "--", DATA_FILE)
        git_net("pull", "--rebase", "--autostash", "origin", branch)
        with open(os.path.join(ROOT, DATA_FILE), encoding="utf-8") as f:
            remote_state = json.load(f)
        save_json(os.path.join(ROOT, DATA_FILE), merge_states(local_state, remote_state))
        if not commit_data(message):
            log("合并后与远端一致，无需推送")
            return
    raise RuntimeError(f"推送连续 {PUSH_RETRIES} 次失败，数据已提交在本地，下次运行会一起推送")


def unpushed_commits() -> int:
    r = git("rev-list", "--count", f"origin/{current_branch()}..HEAD", check=False)
    return int(r.stdout.strip() or 0) if r.returncode == 0 else 0


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
def summarize(path: str) -> None:
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    anns = d.get("announcements", [])
    cutoff = gd_psych._recent_cutoff()
    recent = [a for a in anns if (a.get("publish_date") or "") >= cutoff]
    log(f"数据：公告 {len(anns)} 条（近三个月 {len(recent)} 条），含心理岗 {sum(bool(a.get('has_psych')) for a in anns)} 条，"
        f"打不开 {sum(bool(a.get('error')) for a in anns)} 条，已跳过 {len(d.get('skipped', []))} 条")
    for a in recent:
        flag = "含心理岗" if a.get("has_psych") else ("打不开" if a.get("error") else "无心理岗")
        log(f"    {a.get('publish_date')} {a.get('city')} [{flag}] {a['title'][:50]}")
    bad = [name for name, s in d.get("sites", {}).items() if not s.get("ok")]
    if bad:
        log(f"列表页失败的网站：{'、'.join(bad)}")


def main() -> int:
    global _log_file
    parser = argparse.ArgumentParser(description="本机运行 gd_psych 并推送结果")
    parser.add_argument("--no-push", action="store_true", help="只在本地提交，不推送")
    parser.add_argument("--dry-run", action="store_true", help="只抓取，结果写到 logs/ 下的副本，不动仓库")
    parser.add_argument("--skip-city", action="append", default=[], help="本次不访问的城市，如 --skip-city 珠海")
    parser.add_argument("--only-if-due", action="store_true",
                        help="今天已成功运行过就跳过（计划任务用：三个触发每天只运行一次）")
    args = parser.parse_args()

    os.makedirs(LOG_DIR, exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    _log_file = open(os.path.join(LOG_DIR, f"gd_psych-{today}.log"), "a", encoding="utf-8")
    if args.only_if_due and not is_due(datetime.now()):
        log("今天已成功运行过，本次跳过")
        _log_file.close()
        return 0
    log(f"===== 开始（{'dry-run' if args.dry_run else ('不推送' if args.no_push else '抓取+推送')}）=====")

    try:
        if shutil.which("git") is None:
            raise RuntimeError("找不到 git 命令，请确认 Git 已安装并加入 PATH")

        if not args.dry_run:
            if git("status", "--porcelain", "--", DATA_FILE).stdout.strip():
                raise RuntimeError(f"{DATA_FILE} 有未提交的手动修改，请先处理，避免被覆盖")
            git_net("pull", "--rebase", "--autostash", "origin", current_branch())
            log("已拉取远端最新代码和数据")

        check_gov_reachable()
        log("政府网站直连正常")

        if args.skip_city:
            gd_psych.SITES = [s for s in gd_psych.SITES if s["city"] not in args.skip_city]
            log(f"本次跳过：{'、'.join(args.skip_city)}")
        if args.dry_run:
            copy = os.path.join(LOG_DIR, "gd-psych.dry-run.json")
            shutil.copyfile(os.path.join(ROOT, DATA_FILE), copy)
            gd_psych.STATE_FILE = copy

        records = gd_psych.run(SimpleNamespace(shallow=False))
        log(f"gd_psych 完成，含心理岗公告 {len(records)} 条")
        summarize(gd_psych.STATE_FILE)

        if args.dry_run:
            log(f"dry-run：结果在 {gd_psych.STATE_FILE}，未改动仓库")
            return 0

        message = f"chore(data): 本机更新广东心理岗数据 {today}"
        if commit_data(message):
            log("已提交 gd-psych.json")
        else:
            log("数据没有变化，无需提交")
        if args.no_push:
            log("--no-push：不推送")
        elif unpushed_commits():
            push_with_merge(message)
        if not args.no_push:
            mark_success()  # 没推送的不算完成，下次触发还会再跑
        log("===== 完成 =====")
        return 0
    except Exception as e:  # noqa: BLE001
        log(f"失败：{e}")
        log("===== 结束（失败）=====")
        return 1
    finally:
        _log_file.close()


if __name__ == "__main__":
    sys.exit(main())
