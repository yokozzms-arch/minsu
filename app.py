"""
Samsung Members 커뮤니티 대시보드 - 백엔드 서버 v2
실행: python app.py
접속: http://localhost:5000
"""

from flask import Flask, jsonify, request, send_from_directory, send_file
from flask_cors import CORS
import requests
import time
from datetime import datetime, timezone, timedelta
from collections import defaultdict
import os
import io
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

app = Flask(__name__, static_folder=None)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

# ── 비밀번호 설정 ──────────────────────────────────────────────────────────────
DASHBOARD_PASSWORD = os.environ.get("DASHBOARD_PASSWORD", "samsung2026!")
REPORT_PASSWORD    = os.environ.get("REPORT_PASSWORD", "report2026!")

@app.route("/api/auth/login", methods=["POST"])
def auth_login():
    body = request.json or {}
    pw   = body.get("password", "")
    if pw == DASHBOARD_PASSWORD:
        return jsonify({"status": "ok"})
    return jsonify({"error": "비밀번호가 틀렸어요"}), 401

@app.route("/api/auth/check", methods=["POST"])
def auth_check():
    body = request.json or {}
    pw   = body.get("password", "")
    if pw == DASHBOARD_PASSWORD:
        return jsonify({"status": "ok"})
    return jsonify({"error": "unauthorized"}), 401

@app.route("/api/auth/report", methods=["POST"])
def auth_report():
    body = request.json or {}
    pw   = body.get("password", "")
    if pw == REPORT_PASSWORD:
        return jsonify({"status": "ok"})
    return jsonify({"error": "비밀번호가 틀렸어요"}), 401


@app.after_request
def after_request(response):
    response.headers.add('Access-Control-Allow-Origin', '*')
    response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
    response.headers.add('Access-Control-Allow-Methods', 'GET,PUT,POST,DELETE,OPTIONS')
    response.headers.add('ngrok-skip-browser-warning', 'true')
    return response

BASE_URL   = "https://r1.community.samsung.com"
SEARCH_URL = f"{BASE_URL}/api/2.0/search"
HEADERS    = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8",
}

def liql(query):
    try:
        resp = requests.get(SEARCH_URL, headers=HEADERS, params={"q": query}, timeout=15)
        data = resp.json()
        if data.get("status") == "success":
            return data.get("data", {})
    except:
        pass
    return {}

def liql_count(query):
    return liql(query).get("count", 0)

def parse_dt(dt_str):
    try:
        return datetime.fromisoformat(dt_str)
    except:
        return None

def week_label(dt):
    iso    = dt.isocalendar()
    monday = dt - timedelta(days=dt.weekday())
    sunday = monday + timedelta(days=6)
    return f"{monday.strftime('%Y.%m/%d')}~{sunday.strftime('%m/%d')} ({iso.week}주)"

def week_monday(dt):
    return (dt - timedelta(days=dt.weekday())).strftime("%Y-%m-%d")


# ── API: 유저 프로필 ──────────────────────────────────────────────────────────
@app.route("/api/user/<user_id>")
def get_user(user_id):
    try:
        resp = requests.get(f"{BASE_URL}/api/2.0/users/{user_id}", headers=HEADERS, timeout=15)
        data = resp.json()
        if data.get("status") != "success":
            return jsonify({"error": f"유저를 찾을 수 없습니다 (id: {user_id})"}), 404
        u    = data["data"]
        rank = u.get("rank", {})
        reg  = u.get("registration_data", {})
        metrics = u.get("metrics", {})
        return jsonify({
            "user_id":   user_id,
            "login":     u.get("login", ""),
            "rank":      rank.get("name", "") if isinstance(rank, dict) else "",
            "reg_date":  (reg.get("registration_time") or "")[:10],
            "posts":     metrics.get("posts", 0),
            "avatar":    u.get("avatar", {}).get("profile", "") if isinstance(u.get("avatar"), dict) else "",
            "view_href": u.get("view_href", ""),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── API: 기간 내 게시물 + 주차별 집계 ─────────────────────────────────────────
def collect_period_stats_fast(user_id, dt_from, dt_to):
    """기간 내 포스팅 수 + 좋아요/댓글 합계 빠르게 수집 (게시물별 API 미사용)"""
    # 포스팅 수
    posts = []
    offset, limit = 0, 50
    while True:
        q = (f"SELECT id, post_time "
             f"FROM messages WHERE author.id = '{user_id}' AND depth = 0 "
             f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
        data  = liql(q)
        items = data.get("items", [])
        if not items:
            break
        stop = False
        for item in items:
            dt = parse_dt(item["post_time"])
            if dt is None:
                continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from:
                stop = True
                break
            if dt_utc <= dt_to:
                posts.append(item["id"])
        if stop or len(items) < limit:
            break
        offset += limit
        time.sleep(0.2)

    n = len(posts)
    # 좋아요/댓글은 전체 합산 쿼리로 빠르게
    total_kudos   = liql_count(f"SELECT count(*) FROM kudos WHERE message.author.id = '{user_id}'") if n else 0
    time.sleep(0.2)

    return {
        "posts":       n,
        "kudos":       total_kudos,
        "replies":     0,
        "buzz":        n + total_kudos,
        "avg_kudos":   round(total_kudos / n, 1) if n else 0,
    }


@app.route("/api/user/<user_id>/compare_periods")
def compare_periods(user_id):
    """현재 기간 + 전월 + 전년 동월 자동 비교"""
    date_from = request.args.get("from")
    date_to   = request.args.get("to")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc).replace(day=1)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    period_days = (dt_to - dt_from).days + 1

    # 전월: 같은 일수만큼 이전
    prev_month_to   = dt_from - timedelta(days=1)
    prev_month_from = prev_month_to - timedelta(days=period_days-1)

    # 전년 동월: 1년 전 같은 기간
    try:
        prev_year_from = dt_from.replace(year=dt_from.year - 1)
        prev_year_to   = dt_to.replace(year=dt_to.year - 1)
    except ValueError:
        prev_year_from = dt_from - timedelta(days=365)
        prev_year_to   = dt_to   - timedelta(days=365)

    # 3개 기간 통계 수집
    current    = collect_period_stats_fast(user_id, dt_from, dt_to)
    prev_month = collect_period_stats_fast(user_id, prev_month_from, prev_month_to)
    prev_year  = collect_period_stats_fast(user_id, prev_year_from, prev_year_to)

    def change_rate(curr, prev):
        if prev == 0:
            return None
        return round((curr - prev) / prev * 100, 1)

    return jsonify({
        "user_id": user_id,
        "periods": {
            "current":    {"from": dt_from.strftime("%Y-%m-%d"),       "to": dt_to.strftime("%Y-%m-%d"),          "label": "현재 기간"},
            "prev_month": {"from": prev_month_from.strftime("%Y-%m-%d"),"to": prev_month_to.strftime("%Y-%m-%d"),  "label": "전월"},
            "prev_year":  {"from": prev_year_from.strftime("%Y-%m-%d"), "to": prev_year_to.strftime("%Y-%m-%d"),   "label": "전년 동월"},
        },
        "stats": {
            "current":    current,
            "prev_month": prev_month,
            "prev_year":  prev_year,
        },
        "changes": {
            "vs_prev_month": {k: change_rate(current[k], prev_month[k]) for k in ["posts","kudos","replies","buzz"]},
            "vs_prev_year":  {k: change_rate(current[k], prev_year[k])  for k in ["posts","kudos","replies","buzz"]},
        }
    })


@app.route("/api/user/<user_id>/weekly")
def get_weekly(user_id):
    date_from = request.args.get("from")
    date_to   = request.args.get("to")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식이 올바르지 않습니다 (YYYY-MM-DD)"}), 400

    posts = []
    offset = 0
    limit  = 50
    while True:
        q = (f"SELECT id, subject, post_time, view_href, metrics, board "
             f"FROM messages WHERE author.id = '{user_id}' AND depth = 0 "
             f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
        data  = liql(q)
        items = data.get("items", [])
        if not items:
            break
        stop = False
        for item in items:
            dt = parse_dt(item["post_time"])
            if dt is None:
                continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from:
                stop = True
                break
            if dt_utc <= dt_to:
                posts.append({**item, "_dt": dt_utc})
        if stop or len(items) < limit:
            break
        offset += limit
        time.sleep(0.2)

    all_weeks = {}
    post_list = []
    weekly    = defaultdict(lambda: {"posts": 0, "kudos": 0, "replies": 0, "views": 0})

    for post in posts:
        pid    = post["id"]
        dt     = post["_dt"]
        wl     = week_label(dt)
        wm     = week_monday(dt)
        kudos   = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
        time.sleep(0.15)
        r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
        replies = 0
        if r_data.get('items'):
            replies = r_data['items'][0].get('replies', {}).get('count', 0)
        time.sleep(0.15)

        views = post.get("metrics", {}).get("views", 0) if isinstance(post.get("metrics"), dict) else 0
        weekly[wl]["posts"]   += 1
        weekly[wl]["kudos"]   += kudos
        weekly[wl]["replies"] += replies
        weekly[wl]["views"]   += views
        all_weeks[wl] = wm

        # 게시판 유형
        board_raw = post.get("board", {})
        board_id  = board_raw.get("id", "") if isinstance(board_raw, dict) else ""
        board_map = {
            "kr-community-sp-gs":    "갤럭시 S",
            "kr-community-sp-gfold": "갤럭시 Z",
            "kr-community-wear":     "웨어러블",
            "kr-community-sp-ga":    "갤럭시 A",
            "kr-community-svc-other":"커뮤니티",
            "kr-news":               "뉴스",
        }
        board_name = board_map.get(board_id, board_id or "기타")

        # 제목 유형 패턴 분석
        subject = post["subject"]
        import re
        bracket = re.findall(r"\[([^\]]+)\]", subject)
        if bracket:
            post_type = bracket[0].strip()
            # 유사 유형 통합
            type_map = {"팁후기": "팁", "팁 후기": "팁", "review": "리뷰", "Review": "리뷰",
                        "끄적끄적": "일상", "후기": "후기", "리뷰": "리뷰", "팁": "팁"}
            post_type = type_map.get(post_type, post_type)
        else:
            post_type = "일반"

        post_list.append({
            "id":        pid,
            "subject":   subject,
            "date":      dt.strftime("%Y-%m-%d"),
            "week":      wl,
            "kudos":     kudos,
            "replies":   replies,
            "views":     views,
            "board_id":  board_id,
            "board":     board_name,
            "post_type": post_type,
            "url":       post.get("view_href", f"{BASE_URL}/t5/p/{pid}"),
        })

    sorted_weeks = [wl for wl, _ in sorted(all_weeks.items(), key=lambda x: x[1])]

    # 주차별 증감률 계산
    trends = {}
    for i, w in enumerate(sorted_weeks):
        curr = weekly[w]["kudos"] + weekly[w]["posts"] + weekly[w]["replies"]
        if i > 0:
            prev_w = sorted_weeks[i-1]
            prev   = weekly[prev_w]["kudos"] + weekly[prev_w]["posts"] + weekly[prev_w]["replies"]
            if prev > 0:
                trends[w] = round((curr - prev) / prev * 100, 1)
            else:
                trends[w] = 100.0 if curr > 0 else 0.0
        else:
            trends[w] = None

    # 월별 집계
    monthly = defaultdict(lambda: {"posts": 0, "kudos": 0, "replies": 0, "views": 0})
    for p in post_list:
        month_key = p["date"][:7]  # YYYY-MM
        monthly[month_key]["posts"]   += 1
        monthly[month_key]["kudos"]   += p["kudos"]
        monthly[month_key]["replies"] += p["replies"]
        monthly[month_key]["views"]   += p.get("views", 0)
    sorted_months = sorted(monthly.keys())

    # 월별 증감률
    month_trends = {}
    for i, m in enumerate(sorted_months):
        curr = monthly[m]["kudos"] + monthly[m]["posts"] + monthly[m]["replies"]
        if i > 0:
            prev_m = sorted_months[i-1]
            prev   = monthly[prev_m]["kudos"] + monthly[prev_m]["posts"] + monthly[prev_m]["replies"]
            month_trends[m] = round((curr - prev) / prev * 100, 1) if prev > 0 else (100.0 if curr > 0 else 0.0)
        else:
            month_trends[m] = None

    # 성과 분석
    top_posts   = sorted(post_list, key=lambda x: x["kudos"], reverse=True)[:10]
    avg_kudos   = round(sum(p["kudos"] for p in post_list) / len(post_list), 1) if post_list else 0
    avg_reply   = round(sum(p["replies"] for p in post_list) / len(post_list), 1) if post_list else 0
    avg_views   = round(sum(p["views"] for p in post_list) / len(post_list), 1) if post_list else 0
    engage_rate = round((sum(p["kudos"]+p["replies"] for p in post_list) / len(post_list)), 1) if post_list else 0

    # 게시판별 집계
    board_stats = {}
    for p in post_list:
        b = p["board"]
        if b not in board_stats:
            board_stats[b] = {"posts":0,"kudos":0,"replies":0,"views":0}
        board_stats[b]["posts"]   += 1
        board_stats[b]["kudos"]   += p["kudos"]
        board_stats[b]["replies"] += p["replies"]
        board_stats[b]["views"]   += p["views"]
    for b in board_stats:
        n = board_stats[b]["posts"]
        board_stats[b]["avg_kudos"]  = round(board_stats[b]["kudos"]  / n, 1)
        board_stats[b]["avg_views"]  = round(board_stats[b]["views"]  / n, 1)
        board_stats[b]["engage_rate"]= round((board_stats[b]["kudos"]+board_stats[b]["replies"]) / n, 1)

    # 게시물 유형별 집계
    type_stats = {}
    for p in post_list:
        t = p["post_type"]
        if t not in type_stats:
            type_stats[t] = {"posts":0,"kudos":0,"replies":0,"views":0}
        type_stats[t]["posts"]   += 1
        type_stats[t]["kudos"]   += p["kudos"]
        type_stats[t]["replies"] += p["replies"]
        type_stats[t]["views"]   += p["views"]
    for t in type_stats:
        n = type_stats[t]["posts"]
        type_stats[t]["avg_kudos"]  = round(type_stats[t]["kudos"]  / n, 1)
        type_stats[t]["avg_views"]  = round(type_stats[t]["views"]  / n, 1)
        type_stats[t]["engage_rate"]= round((type_stats[t]["kudos"]+type_stats[t]["replies"]) / n, 1)

    # 인사이트 코멘트 자동 생성
    best_board = max(board_stats, key=lambda x: board_stats[x]["avg_kudos"]) if board_stats else "-"
    best_type  = max(type_stats,  key=lambda x: type_stats[x]["avg_kudos"])  if type_stats  else "-"
    tip_avg    = type_stats.get("팁", {}).get("avg_kudos", 0)
    gen_avg    = type_stats.get("일반", {}).get("avg_kudos", 0)
    tip_effect = round(tip_avg / gen_avg, 1) if gen_avg > 0 and tip_avg > 0 else None

    return jsonify({
        "user_id":       user_id,
        "period":        {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        "total_posts":   sum(v["posts"]   for v in weekly.values()),
        "total_kudos":   sum(v["kudos"]   for v in weekly.values()),
        "total_replies": sum(v["replies"] for v in weekly.values()),
        "total_views":   sum(v["views"]   for v in weekly.values()),
        "avg_kudos":     avg_kudos,
        "avg_reply":     avg_reply,
        "avg_views":     avg_views,
        "engage_rate":   engage_rate,
        "weeks":         sorted_weeks,
        "weekly":        {w: {**weekly[w], "trend": trends[w]} for w in sorted_weeks},
        "months":        sorted_months,
        "monthly":       {m: {**monthly[m], "trend": month_trends[m]} for m in sorted_months},
        "top_posts":     top_posts,
        "posts":         sorted(post_list, key=lambda x: x["date"], reverse=True),
        "board_stats":   board_stats,
        "type_stats":    type_stats,
        "best_board":    best_board,
        "best_type":     best_type,
        "tip_effect":    tip_effect,
    })


# ── API: 댓글 활동 수집 ──────────────────────────────────────────────────────
@app.route("/api/user/<user_id>/comments")
def get_user_comments(user_id):
    date_from = request.args.get("from")
    date_to   = request.args.get("to")
    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    comments = []
    offset, limit = 0, 50
    while True:
        q = (f"SELECT id, subject, post_time, depth, board, view_href "
             f"FROM messages WHERE author.id = '{user_id}' AND depth > 0 "
             f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
        data  = liql(q)
        items = data.get("items", [])
        if not items:
            break
        stop = False
        for item in items:
            dt = parse_dt(item["post_time"])
            if dt is None:
                continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from:
                stop = True
                break
            if dt_utc <= dt_to:
                board_raw  = item.get("board", {})
                board_id   = board_raw.get("id", "") if isinstance(board_raw, dict) else ""
                board_map  = {
                    "kr-community-sp-gs":    "갤럭시 S",
                    "kr-community-sp-gfold": "갤럭시 Z",
                    "kr-community-wear":     "웨어러블",
                    "kr-community-sp-ga":    "갤럭시 A",
                    "kr-community-svc-other":"커뮤니티",
                    "kr-news":               "뉴스",
                }
                # 원문 제목에서 "Re: " 제거
                subject = item.get("subject", "")
                clean_subject = subject.replace("Re: ", "").replace("Re:Re: ", "").strip()

                comments.append({
                    "id":      item["id"],
                    "subject": clean_subject,
                    "date":    dt_utc.strftime("%Y-%m-%d"),
                    "depth":   item.get("depth", 1),
                    "type":    "대댓글" if item.get("depth", 1) > 1 else "댓글",
                    "board":   board_map.get(board_id, board_id or "기타"),
                    "url":     item.get("view_href", ""),
                })
        if stop or len(items) < limit:
            break
        offset += limit
        time.sleep(0.2)

    # 게시판별 집계
    board_count = {}
    for c in comments:
        b = c["board"]
        board_count[b] = board_count.get(b, 0) + 1
    board_count = dict(sorted(board_count.items(), key=lambda x: x[1], reverse=True))

    return jsonify({
        "user_id":      user_id,
        "period":       {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        "total":        len(comments),
        "direct":       sum(1 for c in comments if c["depth"] == 1),
        "replies":      sum(1 for c in comments if c["depth"] > 1),
        "board_count":  board_count,
        "comments":     sorted(comments, key=lambda x: x["date"], reverse=True),
    })


# ── API: 다중 유저 요약 + 인사이트 ────────────────────────────────────────────
@app.route("/api/user/<user_id>/summary", methods=["POST"])
def get_user_summary(user_id):
    """단일 계정 요약 (다중 계정 순차 조회용)"""
    body      = request.json or {}
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    try:
        # 프로필
        resp = requests.get(f"{BASE_URL}/api/2.0/users/{user_id}", headers=HEADERS, timeout=15)
        u    = resp.json().get("data", {})
        rank = u.get("rank", {})
        nickname = u.get("login", user_id)

        # 기간 내 게시물 수집
        posts = []
        offset, limit = 0, 50
        while True:
            q = (f"SELECT id, post_time, metrics FROM messages WHERE author.id = '{user_id}' AND depth = 0 "
                 f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
            data  = liql(q)
            items = data.get("items", [])
            if not items:
                break
            stop = False
            for item in items:
                dt = parse_dt(item["post_time"])
                if dt is None:
                    continue
                dt_utc = dt.astimezone(timezone.utc)
                if dt_utc < dt_from:
                    stop = True
                    break
                if dt_utc <= dt_to:
                    views = item.get("metrics", {}).get("views", 0) if isinstance(item.get("metrics"), dict) else 0
                    posts.append({"id": item["id"], "views": views})
            if stop or len(items) < limit:
                break
            offset += limit
            time.sleep(0.2)

        total_posts = len(posts)
        total_kudos = total_replies = total_views = 0

        for post in posts:
            pid = post["id"]
            k = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
            time.sleep(0.15)
            r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
            r = r_data["items"][0].get("replies", {}).get("count", 0) if r_data.get("items") else 0
            time.sleep(0.15)
            total_kudos   += k
            total_replies += r
            total_views   += post["views"]

        buzz = total_posts + total_kudos + total_replies
        # 댓글 수 빠르게 수집
        comment_count = 0
        c_offset = 0
        while True:
            cq = (f"SELECT id, post_time FROM messages WHERE author.id = '{user_id}' AND depth > 0 "
                  f"ORDER BY post_time DESC LIMIT 50 OFFSET {c_offset}")
            cdata = liql(cq)
            citems = cdata.get("items", [])
            if not citems:
                break
            stop_c = False
            for ci in citems:
                cdt = parse_dt(ci["post_time"])
                if cdt is None:
                    continue
                cdt_utc = cdt.astimezone(timezone.utc)
                if cdt_utc < dt_from:
                    stop_c = True
                    break
                if cdt_utc <= dt_to:
                    comment_count += 1
            if stop_c or len(citems) < 50:
                break
            c_offset += 50
            time.sleep(0.2)

        return jsonify({
            "user_id":      user_id,
            "login":        nickname,
            "rank":         rank.get("name", "") if isinstance(rank, dict) else "",
            "posts":        total_posts,
            "kudos":        total_kudos,
            "replies":      total_replies,
            "views":        total_views,
            "comments_written": comment_count,
            "buzz":         buzz,
            "contribution": 0,
            "engage_rate":  round((total_kudos + total_replies) / total_posts, 1) if total_posts else 0,
        })
    except Exception as e:
        return jsonify({"user_id": user_id, "login": user_id, "error": str(e)}), 500


@app.route("/api/users/summary", methods=["POST"])
def get_users_summary():
    body      = request.json or {}
    user_ids  = body.get("user_ids", [])
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    # 기간 파싱
    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        dt_to   = datetime.now(timezone.utc)

    results = []
    for uid in user_ids:
        try:
            # 프로필
            resp = requests.get(f"{BASE_URL}/api/2.0/users/{uid}", headers=HEADERS, timeout=15)
            u    = resp.json().get("data", {})
            rank = u.get("rank", {})
            nickname = u.get("login", uid)

            # 기간 내 게시물 수집 (단일 계정과 동일 방식)
            posts = []
            offset = 0
            limit  = 50
            while True:
                q = (f"SELECT id, post_time FROM messages WHERE author.id = '{uid}' AND depth = 0 "
                     f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
                data  = liql(q)
                items = data.get("items", [])
                if not items:
                    break
                stop = False
                for item in items:
                    dt = parse_dt(item["post_time"])
                    if dt is None:
                        continue
                    dt_utc = dt.astimezone(timezone.utc)
                    if dt_utc < dt_from:
                        stop = True
                        break
                    if dt_utc <= dt_to:
                        posts.append({**item, "_dt": dt_utc})
                if stop or len(items) < limit:
                    break
                offset += limit
                time.sleep(0.2)

            # 게시물별 좋아요/댓글 집계
            total_kudos   = 0
            total_replies = 0
            for post in posts:
                pid = post["id"]
                k   = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
                time.sleep(0.15)
                r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
                r = 0
                if r_data.get('items'):
                    r = r_data['items'][0].get('replies', {}).get('count', 0)
                time.sleep(0.15)
                total_kudos   += k
                total_replies += r

            total_posts = len(posts)
            results.append({
                "user_id":  uid,
                "login":    nickname,
                "rank":     rank.get("name", "") if isinstance(rank, dict) else "",
                "posts":    total_posts,
                "kudos":    total_kudos,
                "replies":  total_replies,
                "buzz":     total_posts + total_kudos + total_replies,
            })
        except Exception as e:
            results.append({"user_id": uid, "login": uid, "error": str(e)})
        time.sleep(0.3)

    results.sort(key=lambda x: x.get("buzz", 0), reverse=True)

    # 인사이트 생성
    active   = [r for r in results if not r.get("error") and r.get("buzz", 0) > 0]
    inactive = [r for r in results if not r.get("error") and r.get("buzz", 0) == 0]
    total_buzz = sum(r.get("buzz", 0) for r in results)

    # 기여도 %
    for r in results:
        r["contribution"] = round(r.get("buzz", 0) / total_buzz * 100, 1) if total_buzz > 0 else 0

    # MVP: 버즈 1위
    mvp = active[0] if active else None

    # 반응률 (좋아요+댓글 / 포스팅)
    for r in results:
        p = r.get("posts", 0)
        r["engage_rate"] = round((r.get("kudos",0) + r.get("replies",0)) / p, 1) if p > 0 else 0

    insights = {
        "mvp":            mvp["login"] if mvp else "-",
        "mvp_buzz":       mvp["buzz"] if mvp else 0,
        "active_count":   len(active),
        "inactive_count": len(inactive),
        "inactive_names": [r["login"] for r in inactive],
        "total_buzz":     total_buzz,
        "top_engage":     max(active, key=lambda x: x["engage_rate"])["login"] if active else "-",
        "top_engage_rate": max(active, key=lambda x: x["engage_rate"])["engage_rate"] if active else 0,
    }

    return jsonify({"results": results, "insights": insights})


# ── API: 다중 유저 주차별 비교 ────────────────────────────────────────────────
@app.route("/api/users/compare", methods=["POST"])
def compare_users():
    body      = request.json or {}
    user_ids  = body.get("user_ids", [])
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    all_weeks  = {}
    user_data  = {}
    nicknames  = {}

    for uid in user_ids:
        try:
            # 닉네임
            resp = requests.get(f"{BASE_URL}/api/2.0/users/{uid}", headers=HEADERS, timeout=15)
            u    = resp.json().get("data", {})
            nicknames[uid] = u.get("login", uid)
        except:
            nicknames[uid] = uid

        posts = []
        offset = 0
        limit  = 50
        while True:
            q = (f"SELECT id, post_time FROM messages WHERE author.id = '{uid}' AND depth = 0 "
                 f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
            data  = liql(q)
            items = data.get("items", [])
            if not items:
                break
            stop = False
            for item in items:
                dt = parse_dt(item["post_time"])
                if dt is None:
                    continue
                dt_utc = dt.astimezone(timezone.utc)
                if dt_utc < dt_from:
                    stop = True
                    break
                if dt_utc <= dt_to:
                    posts.append({**item, "_dt": dt_utc})
            if stop or len(items) < limit:
                break
            offset += limit
            time.sleep(0.2)

        weekly = defaultdict(lambda: {"posts": 0, "kudos": 0, "replies": 0})
        for post in posts:
            pid  = post["id"]
            dt   = post["_dt"]
            wl   = week_label(dt)
            wm   = week_monday(dt)
            all_weeks[wl] = wm
            kudos   = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
            time.sleep(0.15)
            r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
            replies = 0
            if r_data.get('items'):
                replies = r_data['items'][0].get('replies', {}).get('count', 0)
            time.sleep(0.15)
            weekly[wl]["posts"]   += 1
            weekly[wl]["kudos"]   += kudos
            weekly[wl]["replies"] += replies

        user_data[uid] = dict(weekly)
        time.sleep(0.3)

    sorted_weeks = [wl for wl, _ in sorted(all_weeks.items(), key=lambda x: x[1])]

    return jsonify({
        "weeks":     sorted_weeks,
        "nicknames": nicknames,
        "data":      user_data,
    })


# ── 엑셀 공통 스타일 ──────────────────────────────────────────────────────────
def make_excel(wb):
    """완성된 Workbook을 메모리 버퍼로 변환"""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def header_style():
    fill   = PatternFill("solid", start_color="1428A0")
    font   = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    align  = Alignment(horizontal="center", vertical="center")
    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return fill, font, align, border

def data_style(row_idx):
    alt   = PatternFill("solid", start_color="F0F4FF")
    thin  = Side(style="thin", color="CCCCCC")
    border= Border(left=thin, right=thin, top=thin, bottom=thin)
    fill  = alt if row_idx % 2 == 0 else None
    return fill, border

def set_header(ws, cols_widths):
    fill, font, align, border = header_style()
    for ci, (name, width) in enumerate(cols_widths, start=1):
        cell = ws.cell(row=1, column=ci, value=name)
        cell.fill = fill; cell.font = font
        cell.alignment = align; cell.border = border
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


# ── API: 단일 계정 엑셀 다운로드 ──────────────────────────────────────────────
@app.route("/api/user/<user_id>/excel")
def download_user_excel(user_id):
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl 미설치. pip install openpyxl"}), 500

    date_from = request.args.get("from")
    date_to   = request.args.get("to")
    login     = request.args.get("login", user_id)

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    # 게시물 수집
    posts = []
    offset, limit = 0, 50
    while True:
        q = (f"SELECT id, subject, post_time, view_href, metrics, board "
             f"FROM messages WHERE author.id = '{user_id}' AND depth = 0 "
             f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
        data  = liql(q)
        items = data.get("items", [])
        if not items: break
        stop = False
        for item in items:
            dt = parse_dt(item["post_time"])
            if dt is None: continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from: stop = True; break
            if dt_utc <= dt_to: posts.append({**item, "_dt": dt_utc})
        if stop or len(items) < limit: break
        offset += limit
        time.sleep(0.2)

    weekly = defaultdict(lambda: {"posts": 0, "kudos": 0, "replies": 0})
    post_list = []
    all_weeks = {}
    for post in posts:
        pid  = post["id"]
        dt   = post["_dt"]
        wl   = week_label(dt)
        wm   = week_monday(dt)
        all_weeks[wl] = wm
        kudos  = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
        time.sleep(0.15)
        r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
        replies = r_data["items"][0].get("replies", {}).get("count", 0) if r_data.get("items") else 0
        time.sleep(0.15)
        weekly[wl]["posts"]   += 1
        weekly[wl]["kudos"]   += kudos
        weekly[wl]["replies"] += replies
        views = post.get("metrics", {}).get("views", 0) if isinstance(post.get("metrics"), dict) else 0
        post_list.append({"subject": post["subject"], "date": dt.strftime("%Y-%m-%d"),
                          "week": wl, "kudos": kudos, "replies": replies,
                          "views": views, "url": post.get("view_href", "")})

    sorted_weeks = [wl for wl, _ in sorted(all_weeks.items(), key=lambda x: x[1])]

    wb = openpyxl.Workbook()

    # 시트1: 주차별 요약
    ws1 = wb.active
    ws1.title = "주차별 요약"
    set_header(ws1, [("주차",14),("포스팅",12),("좋아요",12),("댓글",12),("버즈 합계",14)])
    for ri, w in enumerate(sorted_weeks, start=2):
        v = weekly[w]
        buzz = v["posts"]+v["kudos"]+v["replies"]
        fill, border = data_style(ri)
        for ci, val in enumerate([w, v["posts"], v["kudos"], v["replies"], buzz], start=1):
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Arial", size=10)
            if fill: cell.fill = fill
            if isinstance(val, int): cell.number_format = "#,##0"

    # 시트2: 게시물 상세
    ws2 = wb.create_sheet("게시물 상세")
    set_header(ws2, [("제목",50),("주차",12),("작성일",14),("조회수",12),("좋아요",12),("댓글",12),("URL",55)])
    for ri, p in enumerate(sorted(post_list, key=lambda x: x["date"], reverse=True), start=2):
        fill, border = data_style(ri)
        for ci, val in enumerate([p["subject"], p["week"], p["date"], p.get("views",0), p["kudos"], p["replies"], p["url"]], start=1):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (1,6) else "center", vertical="center")
            if fill: cell.fill = fill
            if isinstance(val, int): cell.number_format = "#,##0"

    total_received_comments = sum(p["replies"] for p in post_list)

    # 시트3: 수집 요약
    ws3 = wb.create_sheet("수집 요약")
    summary = [("닉네임", login), ("User ID", user_id),
               ("수집 기간", f"{dt_from.strftime('%Y-%m-%d')} ~ {dt_to.strftime('%Y-%m-%d')}"),
               ("수집 일시", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
               ("총 포스팅", len(post_list)), ("총 좋아요", sum(p["kudos"] for p in post_list)),
               ("받은 댓글", total_received_comments)]
    for i, (k, v) in enumerate(summary, start=1):
        ws3.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        ws3.cell(row=i, column=2, value=v).font = Font(name="Arial")
    ws3.column_dimensions["A"].width = 18
    ws3.column_dimensions["B"].width = 35

    buf = make_excel(wb)
    fname = f"samsung_{login}_{dt_from.strftime('%Y%m%d')}_{dt_to.strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── API: 다중 계정 엑셀 다운로드 ──────────────────────────────────────────────
@app.route("/api/users/excel", methods=["POST"])
def download_users_excel():
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl 미설치. pip install openpyxl"}), 500

    body      = request.json or {}
    results   = body.get("results", [])
    insights  = body.get("insights", {})
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    wb = openpyxl.Workbook()

    # 시트1: 버즈 랭킹
    ws1 = wb.active
    ws1.title = "버즈 랭킹"
    set_header(ws1, [("순위",8),("닉네임",22),("등급",14),("포스팅",12),("조회수",12),
                     ("좋아요",12),("받은댓글",12),("작성댓글",12),("버즈",12),("기여도(%)",12),("반응률",10)])
    medals = {0:"🥇",1:"🥈",2:"🥉"}
    for ri, r in enumerate(results, start=2):
        fill, border = data_style(ri)
        err = r.get("error","")
        vals = [medals.get(ri-2, ri-1), r.get("login",""), r.get("rank",""),
                r.get("posts",0), r.get("views",0), r.get("kudos",0), r.get("replies",0),
                r.get("comments_written",0),
                r.get("buzz",0), r.get("contribution",0), r.get("engage_rate",0)]
        for ci, val in enumerate(vals, start=1):
            cell = ws1.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.font = Font(name="Arial", size=10, color="CC0000" if err else "000000")
            cell.alignment = Alignment(horizontal="left" if ci==2 else "center", vertical="center")
            if fill: cell.fill = fill
            if isinstance(val, (int, float)) and ci > 3: cell.number_format = "#,##0"

    # 시트2: 인사이트 요약
    ws2 = wb.create_sheet("인사이트")
    summary = [("수집 기간", f"{date_from} ~ {date_to}"),
               ("수집 일시", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
               ("총 계정", len(results)), ("활동 계정", insights.get("active_count",0)),
               ("비활성 계정", insights.get("inactive_count",0)),
               ("MVP", insights.get("mvp","-")), ("MVP 버즈", insights.get("mvp_buzz",0)),
               ("최고 반응률", insights.get("top_engage","-")),
               ("총 버즈", insights.get("total_buzz",0)),
               ("비활성 계정 목록", ", ".join(insights.get("inactive_names",[])))]
    for i, (k, v) in enumerate(summary, start=1):
        ws2.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        ws2.cell(row=i, column=2, value=v).font = Font(name="Arial")
    ws2.column_dimensions["A"].width = 18
    ws2.column_dimensions["B"].width = 40

    buf = make_excel(wb)
    fname = f"samsung_multi_{date_from}_{date_to}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── Apify Instagram 수집 ──────────────────────────────────────────────────────
APIFY_API_KEY  = "apify_api_cS8V8n1d5vKVtCcgdn85BJqNdDJvHu1PSo3C"
APIFY_ACTOR    = "apify~instagram-post-scraper"
APIFY_BASE     = "https://api.apify.com/v2"

def apify_headers():
    return {"Authorization": f"Bearer {APIFY_API_KEY}", "Content-Type": "application/json"}

APIFY_PROFILE_ACTOR = "apify~instagram-profile-scraper"

def get_instagram_profile(username):
    """Instagram 프로필 정보 (팔로워/팔로잉/전체 게시물) 조회"""
    try:
        run_resp = requests.post(
            f"{APIFY_BASE}/acts/{APIFY_PROFILE_ACTOR}/runs",
            headers=apify_headers(),
            json={"usernames": [username]},
            timeout=30
        )
        if run_resp.status_code != 201:
            return {}

        run_id     = run_resp.json()["data"]["id"]
        dataset_id = run_resp.json()["data"]["defaultDatasetId"]

        for _ in range(30):
            time.sleep(3)
            status = requests.get(
                f"{APIFY_BASE}/acts/{APIFY_PROFILE_ACTOR}/runs/{run_id}",
                headers=apify_headers(), timeout=15
            ).json()["data"]["status"]
            if status == "SUCCEEDED":
                break
            if status == "FAILED":
                return {}

        items = requests.get(
            f"{APIFY_BASE}/datasets/{dataset_id}/items",
            params={"token": APIFY_API_KEY, "limit": 1},
            timeout=30
        ).json()

        if isinstance(items, list) and items:
            p = items[0]
            return {
                "followers":   p.get("followersCount", 0) or 0,
                "following":   p.get("followsCount",   0) or 0,
                "total_posts": p.get("postsCount",     0) or 0,
                "full_name":   p.get("fullName",       ""),
                "bio":         p.get("biography",      "")[:100],
                "avatar":      p.get("profilePicUrl",  ""),
            }
    except:
        pass
    return {}


# ── Apify 공통 실행/결과 API ──────────────────────────────────────────────────
@app.route("/api/apify/run-sync", methods=["POST"])
def apify_run_sync():
    """Apify Actor 동기 실행 - 결과까지 한번에 반환 (최대 120초)"""
    body       = request.json or {}
    actor      = body.get("actor", APIFY_ACTOR)
    input_data = body.get("input", {})
    timeout    = int(body.get("timeout", 300))
    try:
        url = (f"{APIFY_BASE}/acts/{actor}/run-sync-get-dataset-items"
               f"?token={APIFY_API_KEY}&timeout={timeout}")
        resp = requests.post(
            url,
            headers={"Content-Type": "application/json"},
            json=input_data,
            timeout=timeout + 120
        )
        if resp.status_code not in (200, 201):
            return jsonify({"error": f"Apify 실행 실패 ({resp.status_code}): {resp.text[:200]}"}), 500
        items = resp.json()
        if not isinstance(items, list):
            return jsonify({"error": "결과 형식 오류", "raw": str(items)[:200]}), 500
        return jsonify({"items": items, "count": len(items)})
    except requests.Timeout:
        return jsonify({"error": "Apify 타임아웃 - 계정 게시물이 너무 많거나 Instagram 차단 중"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/apify/run", methods=["POST"])
def apify_run():
    """Apify Actor 실행 후 run_id 즉시 반환 (비동기)"""
    body      = request.json or {}
    actor     = body.get("actor", APIFY_ACTOR)
    input_data = body.get("input", {})
    try:
        run_resp = requests.post(
            f"{APIFY_BASE}/acts/{actor}/runs",
            headers=apify_headers(),
            json=input_data,
            timeout=30
        )
        if run_resp.status_code != 201:
            return jsonify({"error": f"Apify 실행 실패: {run_resp.text}"}), 500
        data = run_resp.json()["data"]
        return jsonify({
            "run_id":     data["id"],
            "dataset_id": data["defaultDatasetId"],
            "status":     data["status"],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/apify/status/<run_id>")
def apify_status(run_id):
    """Apify 실행 상태 확인"""
    actor = request.args.get("actor", APIFY_ACTOR)
    try:
        resp = requests.get(
            f"{APIFY_BASE}/acts/{actor}/runs/{run_id}",
            headers=apify_headers(), timeout=15
        )
        data = resp.json()["data"]
        return jsonify({
            "status":     data["status"],
            "dataset_id": data["defaultDatasetId"],
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/apify/results/<dataset_id>")
def apify_results(dataset_id):
    """Apify 결과 수집"""
    limit = int(request.args.get("limit", 100))
    try:
        resp = requests.get(
            f"{APIFY_BASE}/datasets/{dataset_id}/items",
            params={"token": APIFY_API_KEY, "limit": limit},
            timeout=30
        )
        items = resp.json()
        if not isinstance(items, list):
            return jsonify({"error": "결과 조회 실패", "raw": str(items)[:200]}), 500
        return jsonify({"items": items, "count": len(items)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/apify/instagram/posts", methods=["POST"])
def get_apify_instagram_posts():
    body       = request.json or {}
    username   = body.get("username", "").strip().lstrip("@")
    date_from  = body.get("from", "")
    date_to    = body.get("to", "")
    max_posts  = int(body.get("max_posts", 50))
    dataset_id = body.get("dataset_id", "")  # 이미 실행된 경우 결과만 가져오기

    if not username and not dataset_id:
        return jsonify({"error": "username 필요"}), 400

    try:
        from datetime import timezone as tz
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)

        # dataset_id가 있으면 결과만 가져오기 (이미 실행된 경우)
        if dataset_id:
            items_resp = requests.get(
                f"{APIFY_BASE}/datasets/{dataset_id}/items",
                params={"token": APIFY_API_KEY, "limit": 500},  # 충분히 많이 가져오기
                timeout=30
            )
            all_posts = items_resp.json()
        else:
            # Actor 실행 + 대기 (로컬용, 짧은 타임아웃)
            run_resp = requests.post(
                f"{APIFY_BASE}/acts/{APIFY_ACTOR}/runs",
                headers=apify_headers(),
                json={"username": [username], "resultsLimit": max(max_posts, 100)},
                timeout=30
            )
            if run_resp.status_code != 201:
                return jsonify({"error": f"Apify 실행 실패: {run_resp.text}"}), 500

            run_id     = run_resp.json()["data"]["id"]
            dataset_id = run_resp.json()["data"]["defaultDatasetId"]

            # 완료 대기 (최대 3분)
            for _ in range(36):
                time.sleep(5)
                status_resp = requests.get(
                    f"{APIFY_BASE}/acts/{APIFY_ACTOR}/runs/{run_id}",
                    headers=apify_headers(), timeout=15
                )
                status = status_resp.json()["data"]["status"]
                if status == "SUCCEEDED":
                    break
                if status == "FAILED":
                    return jsonify({"error": "Apify 실행 실패"}), 500

            items_resp = requests.get(
                f"{APIFY_BASE}/datasets/{dataset_id}/items",
                params={"token": APIFY_API_KEY, "limit": max_posts},
                timeout=30
            )
            all_posts = items_resp.json()

        if not isinstance(all_posts, list):
            return jsonify({"error": "결과 조회 실패"}), 500

        # 기간 필터링 + username 필터 (다른 계정 게시물 제거)
        posts = []
        for p in all_posts:
            # username 필터: 요청한 계정의 게시물만
            if username and p.get("ownerUsername", "").lower() != username.lower():
                continue
            try:
                ts = p.get("timestamp", "")
                if not ts:
                    continue
                dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
                dt_utc = dt.astimezone(tz.utc)
                if dt_utc < dt_from or dt_utc > dt_to:
                    continue
            except Exception as te:
                print(f"날짜 파싱 오류: {p.get('timestamp')} → {te}")
                continue

            caption      = (p.get("caption") or "")[:100]
            raw_type     = p.get("type", "")
            product_type = p.get("productType", "") or ""
            # Reels 구분: productType이 clips인 경우
            if product_type.lower() == "clips":
                post_type = "Reels"
            elif raw_type == "Video":
                post_type = "동영상"
            elif raw_type == "Sidecar":
                post_type = "슬라이드"
            elif raw_type == "Image":
                post_type = "이미지"
            else:
                post_type = raw_type or "기타"

            posts.append({
                "id":        p.get("id", ""),
                "date":      dt_utc.strftime("%Y-%m-%d"),
                "type":      post_type,
                "caption":   caption,
                "hashtags":  p.get("hashtags", []),
                "likes":     p.get("likesCount", 0) or 0,
                "comments":  p.get("commentsCount", 0) or 0,
                "views":     p.get("videoViewCount", 0) or 0,
                "url":       p.get("url", ""),
                "thumbnail": p.get("displayUrl", ""),
            })

        posts.sort(key=lambda x: x["date"], reverse=True)

        total_likes    = sum(p["likes"]    for p in posts)
        total_comments = sum(p["comments"] for p in posts)
        total_views    = sum(p["views"]    for p in posts)
        n = len(posts)

        # 유형별 집계
        type_stats = {}
        for p in posts:
            t = p["type"] or "기타"
            if t not in type_stats:
                type_stats[t] = {"posts":0,"likes":0,"comments":0,"views":0}
            type_stats[t]["posts"]    += 1
            type_stats[t]["likes"]    += p["likes"]
            type_stats[t]["comments"] += p["comments"]
            type_stats[t]["views"]    += p["views"]
        for t in type_stats:
            tn = type_stats[t]["posts"]
            type_stats[t]["avg_likes"]    = round(type_stats[t]["likes"]    / tn, 1)
            type_stats[t]["avg_comments"] = round(type_stats[t]["comments"] / tn, 1)

        # 프로필 정보 조회
        profile = get_instagram_profile(username)

        return jsonify({
            "username":      username,
            "followers":     profile.get("followers",   0),
            "following":     profile.get("following",   0),
            "total_posts":   profile.get("total_posts", 0),
            "full_name":     profile.get("full_name",   ""),
            "avatar":        profile.get("avatar",      ""),
            "period":        {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
            "period_posts":  n,
            "total_likes":   total_likes,
            "total_comments":total_comments,
            "total_views":   total_views,
            "avg_likes":     round(total_likes    / n, 1) if n else 0,
            "avg_comments":  round(total_comments / n, 1) if n else 0,
            "type_stats":    type_stats,
            "posts":         posts,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/apify/instagram/multi", methods=["POST"])
def get_apify_instagram_multi():
    body      = request.json or {}
    accounts  = body.get("accounts", [])
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    if not accounts:
        return jsonify({"error": "accounts 필요"}), 400

    try:
        from datetime import timezone as tz
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)

        usernames = [a["instagram"].lstrip("@").split("/")[-1].strip("/")
                     for a in accounts if a.get("instagram")]
        if not usernames:
            return jsonify({"error": "Instagram 계정이 없어요"}), 400

        # 전체 계정 한 번에 실행
        run_resp = requests.post(
            f"{APIFY_BASE}/acts/{APIFY_ACTOR}/runs",
            headers=apify_headers(),
            json={"username": usernames, "resultsLimit": 50},
            timeout=30
        )
        if run_resp.status_code != 201:
            return jsonify({"error": f"Apify 실행 실패: {run_resp.text}"}), 500

        run_id     = run_resp.json()["data"]["id"]
        dataset_id = run_resp.json()["data"]["defaultDatasetId"]

        for _ in range(120):
            time.sleep(5)
            status = requests.get(
                f"{APIFY_BASE}/acts/{APIFY_ACTOR}/runs/{run_id}",
                headers=apify_headers(), timeout=15
            ).json()["data"]["status"]
            if status == "SUCCEEDED":
                break
            if status == "FAILED":
                return jsonify({"error": "Apify 실행 실패"}), 500

        all_posts = requests.get(
            f"{APIFY_BASE}/datasets/{dataset_id}/items",
            params={"token": APIFY_API_KEY, "limit": 1000},
            timeout=30
        ).json()

        # 계정별로 분류
        user_posts = {}
        for p in (all_posts if isinstance(all_posts, list) else []):
            uname = p.get("ownerUsername", "").lower()
            try:
                dt = datetime.fromisoformat(p["timestamp"].replace("Z","+00:00")).astimezone(tz.utc)
                if dt < dt_from or dt > dt_to:
                    continue
            except:
                continue
            if uname not in user_posts:
                user_posts[uname] = []
            user_posts[uname].append({
                "likes":    p.get("likesCount",    0) or 0,
                "comments": p.get("commentsCount", 0) or 0,
                "views":    p.get("videoViewCount",0) or 0,
                "date":     dt.strftime("%Y-%m-%d"),
                "url":      p.get("url",""),
            })

        results = []
        for acc in accounts:
            ig_url   = acc.get("instagram","")
            uname    = ig_url.lstrip("@").split("/")[-1].strip("/").lower()
            posts    = user_posts.get(uname, [])
            n        = len(posts)
            results.append({
                "user_id":       acc.get("user_id",""),
                "login":         acc.get("login",""),
                "ig_username":   uname,
                "ig_url":        ig_url,
                "period_posts":  n,
                "total_likes":   sum(p["likes"]    for p in posts),
                "total_comments":sum(p["comments"] for p in posts),
                "avg_likes":     round(sum(p["likes"]/n for p in posts),1) if n else 0,
            })

        results.sort(key=lambda x: x["total_likes"], reverse=True)
        return jsonify({
            "results": results,
            "period":  {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Instagram 수집 ────────────────────────────────────────────────────────────
INSTA_SESSION_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "insta_session")
_insta_loader = None

def get_insta_loader():
    global _insta_loader
    if _insta_loader is None:
        try:
            import instaloader
            _insta_loader = instaloader.Instaloader(
                download_pictures=False,
                download_videos=False,
                download_video_thumbnails=False,
                download_geotags=False,
                download_comments=False,
                save_metadata=False,
                compress_json=False,
            )
        except ImportError:
            return None
    return _insta_loader

@app.route("/api/instagram/login", methods=["POST"])
def instagram_login():
    body = request.json or {}
    ig_user = body.get("username", "").strip()
    ig_pass = body.get("password", "").strip()
    if not ig_user or not ig_pass:
        return jsonify({"error": "username, password 필요"}), 400
    try:
        import instaloader
        L = get_insta_loader()
        if L is None:
            return jsonify({"error": "instaloader 미설치. pip install instaloader"}), 500
        L.login(ig_user, ig_pass)
        L.save_session_to_file(INSTA_SESSION_FILE)
        # 로그인 계정명 별도 저장
        with open(INSTA_SESSION_FILE + ".user", "w") as f:
            f.write(ig_user)
        return jsonify({"status": "ok", "message": f"{ig_user} 로그인 성공"})
    except instaloader.exceptions.BadCredentialsException:
        return jsonify({"error": "아이디 또는 비밀번호가 틀렸어요"}), 401
    except instaloader.exceptions.TwoFactorAuthRequiredException:
        return jsonify({"error": "2단계 인증이 필요해요. Instagram 앱에서 해제 후 시도해주세요"}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/instagram/status")
def instagram_status():
    if os.path.exists(INSTA_SESSION_FILE):
        user_file = INSTA_SESSION_FILE + ".user"
        username  = ""
        if os.path.exists(user_file):
            with open(user_file) as f:
                username = f.read().strip()
        return jsonify({"logged_in": True, "username": username})
    return jsonify({"logged_in": False})

@app.route("/api/instagram/logout", methods=["POST"])
def instagram_logout():
    global _insta_loader
    if os.path.exists(INSTA_SESSION_FILE):
        os.remove(INSTA_SESSION_FILE)
    if os.path.exists(INSTA_SESSION_FILE + ".user"):
        os.remove(INSTA_SESSION_FILE + ".user")
    _insta_loader = None
    return jsonify({"status": "ok"})

@app.route("/api/instagram/posts/<ig_username>")
def get_instagram_posts(ig_username):
    date_from = request.args.get("from")
    date_to   = request.args.get("to")
    try:
        import instaloader
        from datetime import timezone as tz
        L = get_insta_loader()
        if L is None:
            return jsonify({"error": "instaloader 미설치"}), 500

        # 세션 로드
        if os.path.exists(INSTA_SESSION_FILE):
            try:
                session_user = "stars_2026__"
                user_file = INSTA_SESSION_FILE + ".user"
                if os.path.exists(user_file):
                    with open(user_file) as f:
                        session_user = f.read().strip()
                L.load_session_from_file(session_user, INSTA_SESSION_FILE)
            except Exception as se:
                print(f"세션 로드 실패: {se}")

        # 기간 파싱
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)

        profile = instaloader.Profile.from_username(L.context, ig_username)
        posts   = []
        for post in profile.get_posts():
            post_dt = post.date_utc.replace(tzinfo=tz.utc)
            if post_dt < dt_from:
                break
            if post_dt > dt_to:
                continue
            posts.append({
                "id":        post.shortcode,
                "date":      post_dt.strftime("%Y-%m-%d"),
                "caption":   (post.caption or "")[:100],
                "likes":     post.likes,
                "comments":  post.comments,
                "views":     post.video_view_count if post.is_video else 0,
                "is_video":  post.is_video,
                "type":      "동영상" if post.is_video else "이미지",
                "url":       f"https://www.instagram.com/p/{post.shortcode}/",
                "thumbnail": post.url,
            })
            time.sleep(0.5)

        total_likes    = sum(p["likes"]    for p in posts)
        total_comments = sum(p["comments"] for p in posts)
        avg_likes      = round(total_likes    / len(posts), 1) if posts else 0
        avg_comments   = round(total_comments / len(posts), 1) if posts else 0

        # 유형별 집계
        type_stats = {}
        for p in posts:
            t = p["type"]
            if t not in type_stats:
                type_stats[t] = {"posts": 0, "likes": 0, "comments": 0}
            type_stats[t]["posts"]    += 1
            type_stats[t]["likes"]    += p["likes"]
            type_stats[t]["comments"] += p["comments"]
        for t in type_stats:
            n = type_stats[t]["posts"]
            type_stats[t]["avg_likes"]    = round(type_stats[t]["likes"]    / n, 1)
            type_stats[t]["avg_comments"] = round(type_stats[t]["comments"] / n, 1)

        return jsonify({
            "username":      ig_username,
            "followers":     profile.followers,
            "following":     profile.followees,
            "total_posts":   profile.mediacount,
            "period_posts":  len(posts),
            "period":        {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
            "total_likes":   total_likes,
            "total_comments":total_comments,
            "avg_likes":     avg_likes,
            "avg_comments":  avg_comments,
            "type_stats":    type_stats,
            "posts":         posts,
        })
    except instaloader.exceptions.ProfileNotExistsException:
        return jsonify({"error": f"@{ig_username} 계정을 찾을 수 없어요"}), 404
    except instaloader.exceptions.LoginRequiredException:
        return jsonify({"error": "로그인이 필요해요. Instagram 로그인 탭에서 먼저 로그인해주세요"}), 401
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── YouTube 수집 ──────────────────────────────────────────────────────────────
YOUTUBE_API_KEY = "AIzaSyCCsxPLPLypmDniai5zoslwLXmdNEOFaUM"
YOUTUBE_API_URL = "https://www.googleapis.com/youtube/v3"

def yt_get(endpoint, params):
    params["key"] = YOUTUBE_API_KEY
    resp = requests.get(f"{YOUTUBE_API_URL}/{endpoint}", params=params, timeout=15)
    return resp.json()

def get_channel_id(channel_input):
    """채널 URL, @핸들, 채널ID 등 다양한 형식 처리"""
    # @핸들 형식
    if channel_input.startswith("@") or "youtube.com/@" in channel_input:
        handle = channel_input.split("/@")[-1].split("/")[0]
        if not handle.startswith("@"):
            handle = "@" + handle
        data = yt_get("channels", {"part": "id,snippet", "forHandle": handle})
        items = data.get("items", [])
        if items:
            return items[0]["id"], items[0]["snippet"]["title"]
    # 채널 ID 형식 (UC로 시작)
    if channel_input.startswith("UC") and len(channel_input) == 24:
        data = yt_get("channels", {"part": "id,snippet", "id": channel_input})
        items = data.get("items", [])
        if items:
            return items[0]["id"], items[0]["snippet"]["title"]
    # URL에서 추출
    if "youtube.com/channel/" in channel_input:
        cid = channel_input.split("youtube.com/channel/")[-1].split("/")[0]
        data = yt_get("channels", {"part": "id,snippet", "id": cid})
        items = data.get("items", [])
        if items:
            return items[0]["id"], items[0]["snippet"]["title"]
    # 검색으로 찾기
    data = yt_get("search", {"part": "snippet", "q": channel_input, "type": "channel", "maxResults": 1})
    items = data.get("items", [])
    if items:
        cid = items[0]["snippet"]["channelId"]
        title = items[0]["snippet"]["channelTitle"]
        return cid, title
    return None, None

@app.route("/api/youtube/channel")
def get_youtube_channel():
    channel_input = request.args.get("channel", "").strip()
    if not channel_input:
        return jsonify({"error": "channel 파라미터 필요"}), 400
    try:
        channel_id, channel_title = get_channel_id(channel_input)
        if not channel_id:
            return jsonify({"error": f"채널을 찾을 수 없어요: {channel_input}"}), 404

        data = yt_get("channels", {
            "part": "snippet,statistics,brandingSettings",
            "id": channel_id
        })
        items = data.get("items", [])
        if not items:
            return jsonify({"error": "채널 정보를 가져올 수 없어요"}), 404

        ch = items[0]
        stats = ch.get("statistics", {})
        return jsonify({
            "channel_id":   channel_id,
            "title":        ch["snippet"]["title"],
            "description":  ch["snippet"].get("description", "")[:200],
            "subscribers":  int(stats.get("subscriberCount", 0)),
            "total_views":  int(stats.get("viewCount", 0)),
            "video_count":  int(stats.get("videoCount", 0)),
            "thumbnail":    ch["snippet"].get("thumbnails", {}).get("default", {}).get("url", ""),
            "url":          f"https://www.youtube.com/channel/{channel_id}",
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/youtube/videos")
def get_youtube_videos():
    channel_input = request.args.get("channel", "").strip()
    date_from     = request.args.get("from", "")
    date_to       = request.args.get("to", "")
    if not channel_input:
        return jsonify({"error": "channel 파라미터 필요"}), 400
    try:
        from datetime import timezone as tz
        channel_id, _ = get_channel_id(channel_input)
        if not channel_id:
            return jsonify({"error": f"채널을 찾을 수 없어요: {channel_input}"}), 404

        # 기간 파싱
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)

        # 영상 목록 수집 (Search API)
        videos = []
        next_page = None
        while True:
            params = {
                "part":          "snippet",
                "channelId":     channel_id,
                "type":          "video",
                "order":         "date",
                "maxResults":    50,
                "publishedAfter":  dt_from.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "publishedBefore": dt_to.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            if next_page:
                params["pageToken"] = next_page
            data = yt_get("search", params)
            items = data.get("items", [])
            if not items:
                break
            for item in items:
                videos.append({
                    "video_id": item["id"]["videoId"],
                    "title":    item["snippet"]["title"],
                    "date":     item["snippet"]["publishedAt"][:10],
                    "thumbnail":item["snippet"]["thumbnails"].get("default", {}).get("url", ""),
                })
            next_page = data.get("nextPageToken")
            if not next_page:
                break
            time.sleep(0.2)

        if not videos:
            return jsonify({
                "channel_id": channel_id,
                "period": {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
                "videos": [], "total_views": 0, "total_likes": 0, "total_comments": 0,
                "avg_views": 0, "avg_likes": 0, "type_stats": {}
            })

        # 영상 상세 통계 (Videos API, 50개씩)
        video_details = {}
        for i in range(0, len(videos), 50):
            chunk = videos[i:i+50]
            ids   = ",".join(v["video_id"] for v in chunk)
            data  = yt_get("videos", {"part": "statistics,contentDetails", "id": ids})
            for item in data.get("items", []):
                vid   = item["id"]
                stats = item.get("statistics", {})
                dur   = item.get("contentDetails", {}).get("duration", "")
                # 쇼츠 판별 (1분 이하)
                is_short = False
                if "PT" in dur:
                    import re
                    m = re.search(r"PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?", dur)
                    if m:
                        h = int(m.group(1) or 0)
                        mn = int(m.group(2) or 0)
                        s = int(m.group(3) or 0)
                        total_sec = h*3600 + mn*60 + s
                        is_short = total_sec <= 60
                video_details[vid] = {
                    "views":    int(stats.get("viewCount",    0)),
                    "likes":    int(stats.get("likeCount",    0)),
                    "comments": int(stats.get("commentCount", 0)),
                    "is_short": is_short,
                    "type":     "Shorts" if is_short else "일반 영상",
                }
            time.sleep(0.1)

        # 최종 데이터 합치기
        result_videos = []
        for v in videos:
            d = video_details.get(v["video_id"], {"views":0,"likes":0,"comments":0,"is_short":False,"type":"일반 영상"})
            result_videos.append({
                "video_id": v["video_id"],
                "title":    v["title"],
                "date":     v["date"],
                "views":    d["views"],
                "likes":    d["likes"],
                "comments": d["comments"],
                "type":     d["type"],
                "url":      f"https://www.youtube.com/watch?v={v['video_id']}",
                "thumbnail":v["thumbnail"],
            })

        result_videos.sort(key=lambda x: x["date"], reverse=True)

        total_views    = sum(v["views"]    for v in result_videos)
        total_likes    = sum(v["likes"]    for v in result_videos)
        total_comments = sum(v["comments"] for v in result_videos)
        n = len(result_videos)

        # 유형별 집계
        type_stats = {}
        for v in result_videos:
            t = v["type"]
            if t not in type_stats:
                type_stats[t] = {"videos":0,"views":0,"likes":0,"comments":0}
            type_stats[t]["videos"]   += 1
            type_stats[t]["views"]    += v["views"]
            type_stats[t]["likes"]    += v["likes"]
            type_stats[t]["comments"] += v["comments"]
        for t in type_stats:
            tn = type_stats[t]["videos"]
            type_stats[t]["avg_views"]    = round(type_stats[t]["views"]    / tn)
            type_stats[t]["avg_likes"]    = round(type_stats[t]["likes"]    / tn, 1)
            type_stats[t]["avg_comments"] = round(type_stats[t]["comments"] / tn, 1)

        return jsonify({
            "channel_id":    channel_id,
            "period":        {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
            "period_videos": n,
            "total_views":   total_views,
            "total_likes":   total_likes,
            "total_comments":total_comments,
            "avg_views":     round(total_views    / n) if n else 0,
            "avg_likes":     round(total_likes    / n, 1) if n else 0,
            "avg_comments":  round(total_comments / n, 1) if n else 0,
            "type_stats":    type_stats,
            "videos":        result_videos,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── 블로그 RSS 수집 ────────────────────────────────────────────────────────────
import xml.etree.ElementTree as ET
from email.utils import parsedate_to_datetime

def parse_rss_date(date_str):
    """RSS 날짜 문자열 → datetime"""
    if not date_str:
        return None
    try:
        return parsedate_to_datetime(date_str).replace(tzinfo=timezone.utc)
    except:
        try:
            from datetime import timezone as tz
            for fmt in ["%a, %d %b %Y %H:%M:%S %z", "%a, %d %b %Y %H:%M:%S GMT"]:
                try:
                    return datetime.strptime(date_str.strip(), fmt).replace(tzinfo=timezone.utc)
                except:
                    pass
        except:
            pass
    return None

def fetch_rss(rss_url, dt_from, dt_to, blog_type="naver"):
    """RSS 피드에서 기간 내 게시물 수집"""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    posts = []
    try:
        # 네이버 블로그 RSS URL 변환
        if blog_type == "naver":
            blog_id = rss_url.rstrip("/").split("/")[-1]
            rss_url = f"https://rss.blog.naver.com/{blog_id}.xml"
        elif blog_type == "tistory":
            base = rss_url.rstrip("/").split("/rss")[0]
            rss_url = f"{base}/rss"

        resp = requests.get(rss_url, headers=headers, timeout=15)
        resp.raise_for_status()
        root = ET.fromstring(resp.text)
        items = root.findall('.//item')

        for item in items:
            title   = item.findtext('title', '').strip()
            link    = item.findtext('link', '').strip()
            pubdate = item.findtext('pubDate', '').strip()
            desc    = item.findtext('description', '').strip()

            dt = parse_rss_date(pubdate)
            if dt is None:
                continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from:
                break  # RSS는 최신순이므로 이후는 더 오래됨
            if dt_utc > dt_to:
                continue

            # 태그 제거
            import re
            clean_desc = re.sub(r'<[^>]+>', '', desc)[:100]

            posts.append({
                "title":   title,
                "link":    link,
                "date":    dt_utc.strftime("%Y-%m-%d"),
                "preview": clean_desc,
            })
    except Exception as e:
        return [], str(e)
    return posts, None

@app.route("/api/blog/posts", methods=["POST"])
def get_blog_posts():
    body      = request.json or {}
    user_id   = body.get("user_id", "")
    login     = body.get("login", "")
    naver_url = body.get("naver", "").strip()
    tistory_url = body.get("tistory", "").strip()
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    result = {
        "user_id": user_id,
        "login":   login,
        "period":  {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        "naver":   {"posts": [], "error": None, "url": naver_url},
        "tistory": {"posts": [], "error": None, "url": tistory_url},
    }

    if naver_url:
        posts, err = fetch_rss(naver_url, dt_from, dt_to, "naver")
        result["naver"]["posts"] = posts
        result["naver"]["error"] = err

    if tistory_url:
        posts, err = fetch_rss(tistory_url, dt_from, dt_to, "tistory")
        result["tistory"]["posts"] = posts
        result["tistory"]["error"] = err

    return jsonify(result)

@app.route("/api/blog/multi", methods=["POST"])
def get_blog_multi():
    """SNS 등록된 전체 계정 블로그 일괄 조회"""
    body      = request.json or {}
    accounts  = body.get("accounts", [])
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    results = []
    for acc in accounts:
        naver_url   = acc.get("naver", "").strip()
        tistory_url = acc.get("tistory", "").strip()
        if not naver_url and not tistory_url:
            continue

        naver_posts, naver_err     = fetch_rss(naver_url,   dt_from, dt_to, "naver")   if naver_url   else ([], None)
        tistory_posts, tistory_err = fetch_rss(tistory_url, dt_from, dt_to, "tistory") if tistory_url else ([], None)

        results.append({
            "user_id":       acc.get("user_id", ""),
            "login":         acc.get("login", ""),
            "naver_url":     naver_url,
            "tistory_url":   tistory_url,
            "naver_count":   len(naver_posts),
            "tistory_count": len(tistory_posts),
            "total_count":   len(naver_posts) + len(tistory_posts),
            "naver_posts":   naver_posts,
            "tistory_posts": tistory_posts,
            "naver_error":   naver_err,
            "tistory_error": tistory_err,
        })
        time.sleep(0.3)

    results.sort(key=lambda x: x["total_count"], reverse=True)
    return jsonify({
        "results": results,
        "period":  {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
    })


# ── YouTube 채널 검색 ──────────────────────────────────────────────────────────
@app.route("/api/youtube/search")
def search_youtube_channel():
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "검색어 필요"}), 400
    try:
        data = yt_get("search", {
            "part": "snippet",
            "q": q,
            "type": "channel",
            "maxResults": 5,
        })
        results = []
        for item in data.get("items", []):
            snippet = item["snippet"]
            cid = snippet["channelId"]
            results.append({
                "channel_id":  cid,
                "title":       snippet["channelTitle"],
                "description": snippet.get("description", "")[:80],
                "thumbnail":   snippet.get("thumbnails", {}).get("default", {}).get("url", ""),
                "url":         f"https://www.youtube.com/channel/{cid}",
            })
        return jsonify(results)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── YouTube 다중 채널 일괄 조회 ───────────────────────────────────────────────
@app.route("/api/youtube/multi", methods=["POST"])
def get_youtube_multi():
    body      = request.json or {}
    channels  = body.get("channels", [])  # [{user_id, login, youtube}, ...]
    date_from = body.get("from", "")
    date_to   = body.get("to", "")
    if not channels:
        return jsonify({"error": "channels 필요"}), 400

    from datetime import timezone as tz
    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    results = []
    for ch_info in channels:
        channel_input = ch_info.get("youtube", "").strip()
        login         = ch_info.get("login", "")
        user_id       = ch_info.get("user_id", "")
        if not channel_input:
            continue
        try:
            channel_id, channel_title = get_channel_id(channel_input)
            if not channel_id:
                results.append({"user_id": user_id, "login": login, "youtube": channel_input, "error": "채널을 찾을 수 없어요"})
                continue

            # 채널 기본 정보
            ch_data = yt_get("channels", {"part": "snippet,statistics", "id": channel_id})
            ch_items = ch_data.get("items", [])
            if not ch_items:
                results.append({"user_id": user_id, "login": login, "youtube": channel_input, "error": "채널 정보 없음"})
                continue
            ch = ch_items[0]
            stats = ch.get("statistics", {})

            # 기간 내 영상 수집
            search_data = yt_get("search", {
                "part": "snippet", "channelId": channel_id,
                "type": "video", "order": "date", "maxResults": 50,
                "publishedAfter":  dt_from.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "publishedBefore": dt_to.strftime("%Y-%m-%dT%H:%M:%SZ"),
            })
            video_items = search_data.get("items", [])
            period_videos = len(video_items)

            # 영상 통계
            total_views = total_likes = total_comments = 0
            if video_items:
                ids = ",".join(v["id"]["videoId"] for v in video_items)
                vdata = yt_get("videos", {"part": "statistics", "id": ids})
                for vitem in vdata.get("items", []):
                    vs = vitem.get("statistics", {})
                    total_views    += int(vs.get("viewCount",    0))
                    total_likes    += int(vs.get("likeCount",    0))
                    total_comments += int(vs.get("commentCount", 0))

            results.append({
                "user_id":       user_id,
                "login":         login,
                "channel_title": ch["snippet"]["title"],
                "youtube_url":   f"https://www.youtube.com/channel/{channel_id}",
                "subscribers":   int(stats.get("subscriberCount", 0)),
                "total_views":   int(stats.get("viewCount", 0)),
                "video_count":   int(stats.get("videoCount", 0)),
                "period_videos": period_videos,
                "period_views":  total_views,
                "period_likes":  total_likes,
                "period_comments": total_comments,
                "avg_views":     round(total_views / period_videos) if period_videos else 0,
                "error":         "",
            })
            time.sleep(0.3)
        except Exception as e:
            results.append({"user_id": user_id, "login": login, "youtube": channel_input, "error": str(e)})

    results.sort(key=lambda x: x.get("period_views", 0), reverse=True)
    return jsonify({
        "results": results,
        "period": {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
    })


# ── Apify Instagram 해시태그 수집 ────────────────────────────────────────────
APIFY_HASHTAG_ACTOR = "apify~instagram-hashtag-scraper"

@app.route("/api/apify/instagram/hashtag", methods=["POST"])
def get_hashtag_posts():
    body       = request.json or {}
    hashtags   = body.get("hashtags", [])
    date_from  = body.get("from", "")
    date_to    = body.get("to", "")
    max_posts  = int(body.get("max_posts", 50))

    if not hashtags:
        return jsonify({"error": "hashtags 필요"}), 400

    try:
        from datetime import timezone as tz
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=tz.utc)
        else:
            dt_from = datetime.now(tz.utc) - timedelta(weeks=4)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=tz.utc)
        else:
            dt_to = datetime.now(tz.utc)

        # Actor 실행
        run_resp = requests.post(
            f"{APIFY_BASE}/acts/{APIFY_HASHTAG_ACTOR}/runs",
            headers=apify_headers(),
            json={"hashtags": hashtags, "resultsLimit": max_posts},
            timeout=30
        )
        if run_resp.status_code != 201:
            return jsonify({"error": f"Apify 실행 실패: {run_resp.text}"}), 500

        run_id     = run_resp.json()["data"]["id"]
        dataset_id = run_resp.json()["data"]["defaultDatasetId"]

        # 완료 대기
        for _ in range(60):
            time.sleep(5)
            status = requests.get(
                f"{APIFY_BASE}/acts/{APIFY_HASHTAG_ACTOR}/runs/{run_id}",
                headers=apify_headers(), timeout=15
            ).json()["data"]["status"]
            if status == "SUCCEEDED":
                break
            if status == "FAILED":
                return jsonify({"error": "Apify 실행 실패"}), 500

        # 결과 수집
        all_posts = requests.get(
            f"{APIFY_BASE}/datasets/{dataset_id}/items",
            params={"token": APIFY_API_KEY, "limit": max_posts},
            timeout=30
        ).json()

        if not isinstance(all_posts, list):
            return jsonify({"error": "결과 조회 실패"}), 500

        # 기간 필터링 및 정리
        posts = []
        for p in all_posts:
            try:
                dt = datetime.fromisoformat(p["timestamp"].replace("Z", "+00:00")).astimezone(tz.utc)
                if dt < dt_from or dt > dt_to:
                    continue
            except:
                continue

            raw_type2     = p.get("type", "")
            product_type2 = p.get("productType", "") or ""
            if product_type2.lower() == "clips":
                post_type2 = "Reels"
            elif raw_type2 == "Video":
                post_type2 = "동영상"
            elif raw_type2 == "Sidecar":
                post_type2 = "슬라이드"
            elif raw_type2 == "Image":
                post_type2 = "이미지"
            else:
                post_type2 = raw_type2 or "기타"

            posts.append({
                "id":          p.get("id", ""),
                "date":        dt.strftime("%Y-%m-%d"),
                "type":        post_type2,
                "caption":     (p.get("caption") or "")[:100],
                "hashtags":    p.get("hashtags", []),
                "likes":       p.get("likesCount", 0) or 0,
                "comments":    p.get("commentsCount", 0) or 0,
                "owner":       p.get("ownerUsername", ""),
                "url":         p.get("url", ""),
                "thumbnail":   p.get("displayUrl", ""),
            })

        posts.sort(key=lambda x: x["date"], reverse=True)

        total_likes    = sum(p["likes"]    for p in posts)
        total_comments = sum(p["comments"] for p in posts)
        n = len(posts)

        # 계정별 집계
        owner_stats = {}
        for p in posts:
            o = p["owner"] or "unknown"
            if o not in owner_stats:
                owner_stats[o] = {"posts": 0, "likes": 0, "comments": 0}
            owner_stats[o]["posts"]    += 1
            owner_stats[o]["likes"]    += p["likes"]
            owner_stats[o]["comments"] += p["comments"]
        owner_stats = dict(sorted(owner_stats.items(), key=lambda x: x[1]["likes"], reverse=True))

        # 날짜별 집계
        date_stats = {}
        for p in posts:
            d = p["date"]
            if d not in date_stats:
                date_stats[d] = {"posts": 0, "likes": 0, "comments": 0}
            date_stats[d]["posts"]    += 1
            date_stats[d]["likes"]    += p["likes"]
            date_stats[d]["comments"] += p["comments"]

        return jsonify({
            "hashtags":      hashtags,
            "period":        {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
            "total_posts":   n,
            "total_likes":   total_likes,
            "total_comments":total_comments,
            "avg_likes":     round(total_likes    / n, 1) if n else 0,
            "avg_comments":  round(total_comments / n, 1) if n else 0,
            "owner_stats":   owner_stats,
            "date_stats":    date_stats,
            "posts":         posts,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── SNS 정보 관리 (JSON 파일 기반) ───────────────────────────────────────────
import json as _json

SNS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sns_data.json")

def load_sns():
    if os.path.exists(SNS_FILE):
        with open(SNS_FILE, encoding="utf-8") as f:
            return _json.load(f)
    return {}

def save_sns(data):
    with open(SNS_FILE, "w", encoding="utf-8") as f:
        _json.dump(data, f, ensure_ascii=False, indent=2)

@app.route("/api/sns", methods=["GET"])
def get_sns():
    return jsonify(load_sns())

@app.route("/api/sns", methods=["POST"])
def set_sns():
    body = request.json or {}
    uid  = str(body.get("user_id", "")).strip()
    if not uid:
        return jsonify({"error": "user_id 필요"}), 400
    data = load_sns()
    if uid not in data:
        data[uid] = {}
    data[uid]["instagram"]  = body.get("instagram", "").strip()
    data[uid]["instagrams"] = [x.strip() for x in body.get("instagrams", []) if x.strip()]
    data[uid]["youtube"]   = body.get("youtube",   "").strip()
    data[uid]["naver"]     = body.get("naver",     "").strip()
    data[uid]["tistory"]   = body.get("tistory",   "").strip()
    data[uid]["brunch"]    = body.get("brunch",    "").strip()
    data[uid]["tiktok"]    = body.get("tiktok",    "").strip()
    data[uid]["twitch"]    = body.get("twitch",    "").strip()
    data[uid]["etc1_name"] = body.get("etc1_name", "").strip()
    data[uid]["etc1_url"]  = body.get("etc1_url",  "").strip()
    data[uid]["etc2_name"] = body.get("etc2_name", "").strip()
    data[uid]["etc2_url"]  = body.get("etc2_url",  "").strip()
    data[uid]["memo"]      = body.get("memo",      "").strip()
    if body.get("login"):
        data[uid]["login"] = body.get("login", "").strip()
    save_sns(data)
    return jsonify({"status": "ok", "user_id": uid, "data": data[uid]})

@app.route("/api/sns/<user_id>", methods=["DELETE"])
def delete_sns(user_id):
    data = load_sns()
    if user_id in data:
        del data[user_id]
        save_sns(data)
    return jsonify({"status": "ok"})


# ── 소셜 엑셀 다운로드 ────────────────────────────────────────────────────────
@app.route("/api/excel/instagram", methods=["POST"])
def excel_instagram():
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl 미설치"}), 500
    body = request.json or {}
    username  = body.get("username", "")
    profile   = body.get("profile", {})
    posts     = body.get("posts", [])
    period    = body.get("period", {})

    wb = openpyxl.Workbook()

    # 시트1: 프로필 요약
    ws1 = wb.active
    ws1.title = "프로필 요약"
    summary = [
        ("계정명",      "@" + username),
        ("수집 기간",   f"{period.get('from','')} ~ {period.get('to','')}"),
        ("수집 일시",   datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("팔로워",      profile.get("followers", 0)),
        ("팔로잉",      profile.get("following", 0)),
        ("전체 게시물", profile.get("total_posts", 0)),
        ("기간 게시물", len(posts)),
        ("총 좋아요",   sum(p.get("likes",0) for p in posts)),
        ("총 댓글",     sum(p.get("comments",0) for p in posts)),
        ("평균 좋아요", round(sum(p.get("likes",0) for p in posts)/len(posts),1) if posts else 0),
    ]
    for i, (k, v) in enumerate(summary, start=1):
        ws1.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        ws1.cell(row=i, column=2, value=v).font = Font(name="Arial")
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 35

    # 시트2: 게시물 목록
    ws2 = wb.create_sheet("게시물 목록")
    set_header(ws2, [("날짜",14),("유형",10),("내용 미리보기",50),("해시태그",40),("좋아요",12),("댓글",12),("링크",50)])
    for ri, p in enumerate(sorted(posts, key=lambda x: x.get("date",""), reverse=True), start=2):
        fill, border = data_style(ri)
        hashtags = " ".join(["#"+h for h in p.get("hashtags", [])[:5]])
        for ci, val in enumerate([
            p.get("date",""), p.get("type",""), p.get("caption",""),
            hashtags, p.get("likes",0), p.get("comments",0), p.get("url","")
        ], start=1):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (3,4,7) else "center", vertical="center")
            if fill: cell.fill = fill
            if isinstance(val, int): cell.number_format = "#,##0"

    buf = make_excel(wb)
    fname = f"instagram_{username}_{period.get('from','')}_{period.get('to','')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/excel/youtube", methods=["POST"])
def excel_youtube():
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl 미설치"}), 500
    body    = request.json or {}
    channel = body.get("channel", {})
    videos  = body.get("videos", [])
    period  = body.get("period", {})
    type_stats = body.get("type_stats", {})

    wb = openpyxl.Workbook()

    # 시트1: 채널 요약
    ws1 = wb.active
    ws1.title = "채널 요약"
    summary = [
        ("채널명",      channel.get("title", "")),
        ("수집 기간",   f"{period.get('from','')} ~ {period.get('to','')}"),
        ("수집 일시",   datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("구독자",      channel.get("subscribers", 0)),
        ("전체 조회수", channel.get("total_views", 0)),
        ("전체 영상",   channel.get("video_count", 0)),
        ("기간 영상",   len(videos)),
        ("기간 조회수", sum(v.get("views",0) for v in videos)),
        ("기간 좋아요", sum(v.get("likes",0) for v in videos)),
        ("기간 댓글",   sum(v.get("comments",0) for v in videos)),
        ("평균 조회수", round(sum(v.get("views",0) for v in videos)/len(videos)) if videos else 0),
    ]
    for i, (k, v) in enumerate(summary, start=1):
        ws1.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        ws1.cell(row=i, column=2, value=v).font = Font(name="Arial")
        if isinstance(v, int): ws1.cell(row=i, column=2).number_format = "#,##0"
    ws1.column_dimensions["A"].width = 16
    ws1.column_dimensions["B"].width = 35

    # 시트2: 영상 목록
    ws2 = wb.create_sheet("영상 목록")
    set_header(ws2, [("날짜",14),("유형",12),("제목",50),("👁 조회수",14),("👍 좋아요",12),("💬 댓글",12),("링크",55)])
    for ri, v in enumerate(sorted(videos, key=lambda x: x.get("date",""), reverse=True), start=2):
        fill, border = data_style(ri)
        for ci, val in enumerate([
            v.get("date",""), v.get("type",""), v.get("title",""),
            v.get("views",0), v.get("likes",0), v.get("comments",0), v.get("url","")
        ], start=1):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.border = border
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="left" if ci in (3,7) else "center", vertical="center")
            if fill: cell.fill = fill
            if isinstance(val, int): cell.number_format = "#,##0"

    # 시트3: 유형별 성과
    if type_stats:
        ws3 = wb.create_sheet("유형별 성과")
        set_header(ws3, [("유형",14),("영상 수",12),("총 조회수",14),("평균 조회수",14),("총 좋아요",12),("평균 좋아요",12),("평균 댓글",12)])
        for ri, (t, s) in enumerate(type_stats.items(), start=2):
            fill, border = data_style(ri)
            for ci, val in enumerate([t, s.get("videos",0), s.get("views",0), s.get("avg_views",0),
                                       s.get("likes",0), s.get("avg_likes",0), s.get("avg_comments",0)], start=1):
                cell = ws3.cell(row=ri, column=ci, value=val)
                cell.border = border
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if fill: cell.fill = fill

    buf = make_excel(wb)
    fname = f"youtube_{channel.get('title','channel')}_{period.get('from','')}_{period.get('to','')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/api/excel/blog", methods=["POST"])
def excel_blog():
    if not EXCEL_OK:
        return jsonify({"error": "openpyxl 미설치"}), 500
    body   = request.json or {}
    login  = body.get("login", "")
    naver  = body.get("naver", {})
    tistory = body.get("tistory", {})
    period = body.get("period", {})

    wb = openpyxl.Workbook()

    def make_blog_sheet(wb, title, posts, is_first=False):
        ws = wb.active if is_first else wb.create_sheet(title)
        ws.title = title
        set_header(ws, [("날짜",14),("제목",55),("내용 미리보기",50),("링크",55)])
        for ri, p in enumerate(sorted(posts, key=lambda x: x.get("date",""), reverse=True), start=2):
            fill, border = data_style(ri)
            for ci, val in enumerate([p.get("date",""), p.get("title",""), p.get("preview",""), p.get("link","")], start=1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.border = border
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(horizontal="left" if ci in (2,3,4) else "center", vertical="center")
                if fill: cell.fill = fill

    naver_posts   = naver.get("posts", [])
    tistory_posts = tistory.get("posts", [])

    if naver_posts:
        make_blog_sheet(wb, "네이버 블로그", naver_posts, is_first=True)
    if tistory_posts:
        make_blog_sheet(wb, "티스토리", tistory_posts, is_first=not naver_posts)

    # 요약 시트
    ws_sum = wb.create_sheet("수집 요약")
    summary = [
        ("닉네임",          login),
        ("수집 기간",       f"{period.get('from','')} ~ {period.get('to','')}"),
        ("수집 일시",       datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("네이버 게시물",   len(naver_posts)),
        ("티스토리 게시물", len(tistory_posts)),
        ("총 게시물",       len(naver_posts) + len(tistory_posts)),
    ]
    for i, (k, v) in enumerate(summary, start=1):
        ws_sum.cell(row=i, column=1, value=k).font = Font(bold=True, name="Arial")
        ws_sum.cell(row=i, column=2, value=v).font = Font(name="Arial")
    ws_sum.column_dimensions["A"].width = 18
    ws_sum.column_dimensions["B"].width = 35

    buf = make_excel(wb)
    fname = f"blog_{login}_{period.get('from','')}_{period.get('to','')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ── PPT 리포트 생성 (python-pptx) ────────────────────────────────────────────
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    PPT_OK = True
except ImportError:
    PPT_OK = False

def rgb(hex_str):
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def add_rect(slide, x, y, w, h, fill_hex=None, line_hex=None):
    from pptx.util import Pt
    shape = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill_hex:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill_hex)
    else:
        shape.fill.background()
    if line_hex:
        shape.line.color.rgb = rgb(line_hex)
        shape.line.width = Pt(0.5)
    else:
        shape.line.fill.background()
    return shape

def add_text(slide, text, x, y, w, h, size=12, bold=False, color='000000', align='left', wrap=True):
    from pptx.util import Pt
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.word_wrap = wrap
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if align=='center' else PP_ALIGN.LEFT
    run = p.add_run()
    run.text = str(text)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box

def fmt_n(n):
    try: return f"{int(n):,}"
    except: return str(n) if n else '-'

@app.route("/api/report/ppt", methods=["POST"])
def generate_ppt_report():
    if not PPT_OK:
        return jsonify({"error": "python-pptx 미설치. pip install python-pptx"}), 500

    body    = request.json or {}
    period  = body.get("period", {})
    comm    = body.get("community_results", [])
    insta   = body.get("instagram_results", [])
    yt      = body.get("youtube_results", [])
    blog    = body.get("blog_results", [])
    accounts= body.get("accounts", [])

    C = {
        "primary":   "1428A0",
        "secondary": "2D4ED8",
        "white":     "FFFFFF",
        "light":     "F0F4FF",
        "muted":     "64748B",
        "green":     "10B981",
        "pink":      "E1306C",
        "ytred":     "FF0000",
        "bloggreen": "03C75A",
        "dark":      "0F172A",
    }

    try:
        prs = Presentation()
        prs.slide_width  = Inches(10)
        prs.slide_height = Inches(5.625)
        blank = prs.slide_layouts[6]

        def new_slide():
            return prs.slides.add_slide(blank)

        # ── 슬라이드 1: 표지 ──────────────────────────────────
        s = new_slide()
        add_rect(s, 0, 0, 10, 5.625, C["primary"])
        add_text(s, "Samsung Members Stars", 0.6, 1.5, 8, 1.2, size=36, bold=True, color=C["white"], align='left')
        add_text(s, "월별 활동 리포트", 0.6, 2.7, 6, 0.6, size=18, color=C["white"], align='left')
        period_str = f"{period.get('from','')} ~ {period.get('to','')}"
        add_text(s, period_str, 0.6, 3.4, 6, 0.5, size=13, color=C["white"], align='left')
        add_text(s, f"총 {len(accounts)}개 계정", 0.6, 4.0, 4, 0.4, size=11, color=C["white"], align='left')

        # ── 슬라이드 2: 전체 요약 ─────────────────────────────
        s = new_slide()
        add_rect(s, 0, 0, 10, 0.65, C["primary"])
        add_text(s, "📊  전체 활동 요약", 0.4, 0.08, 9, 0.5, size=18, bold=True, color=C["white"])

        total_buzz  = sum(r.get("buzz",0)  for r in comm)
        total_posts = sum(r.get("posts",0) for r in comm)
        total_kudos = sum(r.get("kudos",0) for r in comm)
        ig_likes    = sum(r.get("total_likes",0)    for r in insta)
        yt_views    = sum(r.get("period_views",0)   for r in yt)
        blog_posts  = sum(r.get("total_count",0)    for r in blog)

        kpis = [
            ("커뮤니티 총 버즈", fmt_n(total_buzz),  C["primary"]),
            ("커뮤니티 포스팅", fmt_n(total_posts),  C["secondary"]),
            ("커뮤니티 좋아요", fmt_n(total_kudos),  C["green"]),
            ("Instagram 좋아요", fmt_n(ig_likes),    C["pink"]),
            ("YouTube 조회수",  fmt_n(yt_views),     C["ytred"]),
            ("블로그 게시물",   fmt_n(blog_posts),   C["bloggreen"]),
        ]
        cols, cw, ch = 3, 2.9, 1.35
        for i, (label, val, col) in enumerate(kpis):
            cx = 0.35 + (i%cols)*(cw+0.2)
            cy = 0.85 + (i//cols)*(ch+0.25)
            add_rect(s, cx, cy, cw, ch, C["white"], "E2E8F0")
            add_rect(s, cx, cy, cw, 0.06, col)
            add_text(s, label, cx+0.1, cy+0.12, cw-0.2, 0.3, size=9,  color=C["muted"])
            add_text(s, val,   cx+0.1, cy+0.42, cw-0.2, 0.55, size=22, bold=True, color=col)

        # ── 슬라이드 3~6: 랭킹 ───────────────────────────────
        ranking_slides = [
            ("🏆  삼성 멤버스 커뮤니티 랭킹", comm,  ["포스팅","좋아요","댓글","버즈"],
             lambda r: [r.get("posts",0),r.get("kudos",0),r.get("replies",0),r.get("buzz",0)],
             lambda r: r.get("login","")),
            ("📷  Instagram 랭킹", insta, ["게시물","좋아요","댓글","조회수"],
             lambda r: [r.get("period_posts",0),r.get("total_likes",0),r.get("total_comments",0),r.get("total_views",0)],
             lambda r: r.get("login","")),
            ("▶️  YouTube 랭킹", yt, ["영상","조회수","좋아요","댓글"],
             lambda r: [r.get("period_videos",0),r.get("period_views",0),r.get("period_likes",0),r.get("period_comments",0)],
             lambda r: r.get("login",r.get("channel_title",""))),
            ("📝  블로그 랭킹", blog, ["네이버","티스토리","합계"],
             lambda r: [r.get("naver_count",0),r.get("tistory_count",0),r.get("total_count",0)],
             lambda r: r.get("login","")),
        ]

        medals = ["🥇","🥈","🥉"]
        for title, data, headers, vals_fn, name_fn in ranking_slides:
            s = new_slide()
            add_rect(s, 0, 0, 10, 0.65, C["primary"])
            add_text(s, title, 0.4, 0.08, 9, 0.5, size=18, bold=True, color=C["white"])

            sorted_data = sorted(data, key=lambda r: (vals_fn(r)[-1] or 0), reverse=True)[:12]
            if not sorted_data:
                add_text(s, "데이터 없음", 1, 2, 8, 1, size=14, color=C["muted"], align='center')
                continue

            for j, h in enumerate(headers):
                add_text(s, h, 3.5+j*1.5, 0.72, 1.4, 0.28, size=9, bold=True, color=C["muted"], align='center')

            for i, r in enumerate(sorted_data):
                y   = 1.02 + i*0.37
                bg  = C["light"] if i < 3 else C["white"]
                add_rect(s, 0.3, y, 9.4, 0.36, bg, "E2E8F0")
                medal = medals[i] if i < 3 else str(i+1)
                add_text(s, medal,      0.35, y+0.04, 0.5,  0.28, size=11 if i>=3 else 14, color=C["muted"], align='center')
                add_text(s, name_fn(r), 0.9,  y+0.06, 2.5,  0.24, size=11, bold=i<3, color=C["dark"])
                for k, v in enumerate(vals_fn(r)):
                    add_text(s, fmt_n(v), 3.5+k*1.5, y+0.06, 1.4, 0.24, size=10, color=C["dark"], align='center')

        # ── 슬라이드 7~: 계정별 상세 ─────────────────────────
        for idx, acc in enumerate(accounts):
            s     = new_slide()
            login = acc.get("login", acc.get("user_id",""))
            add_rect(s, 0, 0, 10, 0.75, C["primary"])
            add_text(s, login, 0.4, 0.1, 7, 0.55, size=20, bold=True, color=C["white"])
            add_text(s, f"{idx+1} / {len(accounts)}", 8.5, 0.15, 1.3, 0.4, size=11, color=C["white"], align='center')

            sections = [
                ("🏆 커뮤니티", acc.get("community",{}), C["primary"],
                 [("포스팅",  "posts"),("좋아요","kudos"),("댓글","replies"),("버즈","buzz")], 0.3, 0.85),
                ("📷 Instagram", acc.get("instagram",{}), C["pink"],
                 [("게시물","period_posts"),("좋아요","total_likes"),("댓글","total_comments"),("조회수","total_views")], 5.0, 0.85),
                ("▶️ YouTube",   acc.get("youtube",{}),   C["ytred"],
                 [("영상","period_videos"),("조회수","period_views"),("좋아요","period_likes"),("구독자","subscribers")], 0.3, 3.2),
                ("📝 블로그",    acc.get("blog",{}),       C["bloggreen"],
                 [("네이버","naver_count"),("티스토리","tistory_count"),("합계","total_count")], 5.0, 3.2),
            ]

            for sec_title, sec_data, col, fields, sx, sy in sections:
                sw = 4.4 if sx < 1 else 4.7
                sh = 2.2 if sy < 1 else 2.1
                add_rect(s, sx, sy, sw, sh, C["white"], "E2E8F0")
                add_rect(s, sx, sy, sw, 0.06, col)
                add_text(s, sec_title, sx+0.1, sy+0.1, sw-0.2, 0.3, size=11, bold=True, color=col)
                for fi, (label, key) in enumerate(fields[:4]):
                    fc  = fi % 2
                    fr  = fi // 2
                    fx  = sx + 0.15 + fc * (sw/2)
                    fy  = sy + 0.45 + fr * 0.72
                    add_text(s, label,               fx, fy,      sw/2-0.1, 0.22, size=9,  color=C["muted"])
                    add_text(s, fmt_n(sec_data.get(key,0)), fx, fy+0.2, sw/2-0.1, 0.38, size=18, bold=True, color=col)

        # ── 마지막 슬라이드 ───────────────────────────────────
        s = new_slide()
        add_rect(s, 0, 0, 10, 5.625, C["primary"])
        add_text(s, "Samsung Members Stars", 0.6, 1.5, 8, 1.2, size=36, bold=True, color=C["white"])
        add_text(s, f"{period_str}  리포트 완료", 0.6, 2.8, 8, 0.6, size=16, color=C["white"])
        add_text(s, "© Samsung Electronics Co., Ltd.", 0, 5.1, 10, 0.4, size=10, color=C["white"], align='center')

        # 파일 저장
        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)

        month = period.get('from', 'report')[:7]
        fname = f"samsung_stars_{month}_report.pptx"
        return send_file(buf, as_attachment=True, download_name=fname,
                         mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

    except Exception as e:
        import traceback
        print("PPT error:", traceback.format_exc())
        return jsonify({"error": str(e)}), 500


# ── 월별 통계 비교 API ────────────────────────────────────────────────────────────
@app.route("/api/users/monthly", methods=["POST"])
def get_users_monthly():
    """전체 스타즈 월별 포스팅/좋아요/댓글 비교"""
    body      = request.json or {}
    user_ids  = body.get("user_ids", [])
    date_from = body.get("from", "")
    date_to   = body.get("to", "")

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(days=180)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    # 월 목록 생성
    months = []
    cur = dt_from.replace(day=1)
    while cur <= dt_to:
        months.append(cur.strftime("%Y-%m"))
        if cur.month == 12:
            cur = cur.replace(year=cur.year+1, month=1)
        else:
            cur = cur.replace(month=cur.month+1)

    results = []
    for uid in user_ids:
        # 프로필
        try:
            resp = requests.get(f"{BASE_URL}/api/2.0/users/{uid}", headers=HEADERS, timeout=10)
            u    = resp.json().get("data", {})
            login = u.get("login", uid)
        except:
            login = uid

        # 게시물 수집
        monthly = defaultdict(lambda: {"posts": 0, "kudos": 0, "replies": 0})
        offset, limit = 0, 50
        while True:
            q = (f"SELECT id, post_time FROM messages WHERE author.id = '{uid}' AND depth = 0 "
                 f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
            data  = liql(q)
            items = data.get("items", [])
            if not items: break
            stop = False
            for item in items:
                dt = parse_dt(item["post_time"])
                if dt is None: continue
                dt_utc = dt.astimezone(timezone.utc)
                if dt_utc < dt_from:
                    stop = True
                    break
                if dt_utc <= dt_to:
                    month_key = dt_utc.strftime("%Y-%m")
                    monthly[month_key]["posts"] += 1
                    # 좋아요/댓글은 빠른 집계
                    pid = item["id"]
                    k = liql_count(f"SELECT count(*) FROM kudos WHERE message.id = '{pid}'")
                    r_data = liql(f"SELECT id, replies.count(*) FROM messages WHERE id = '{pid}'")
                    r = r_data["items"][0].get("replies", {}).get("count", 0) if r_data.get("items") else 0
                    monthly[month_key]["kudos"]   += k
                    monthly[month_key]["replies"] += r
                    time.sleep(0.1)
            if stop or len(items) < limit: break
            offset += limit
            time.sleep(0.2)

        results.append({
            "user_id": uid,
            "login":   login,
            "monthly": {m: monthly.get(m, {"posts":0,"kudos":0,"replies":0}) for m in months},
        })

    return jsonify({
        "months":  months,
        "results": results,
    })


# ── 커뮤니티 레벨별 분석 ─────────────────────────────────────────────────────────
KR_BOARDS_FOR_RANK = [
    "kr-community-sp-gs",
    "kr-community-sp-gfold",
    "kr-community-wear",
    "kr-community-sp-ga",
    "kr-community-other",
    "kr-community-promo",
    "kr-community-pc",
    "kr-community-home",
    "kr-news",
]

@app.route("/api/community/rank_stats", methods=["POST"])
def community_rank_stats():
    """게시판별 스캔으로 특정 기간 레벨별 집계 (날짜 필터 우회)"""
    body      = request.json or {}
    date_from = body.get("from", "")
    date_to   = body.get("to", "")
    max_scan  = int(body.get("max_scan", 50000))
    board_id  = body.get("board_id", "")  # 특정 게시판만

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(days=30)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    target_month = dt_from.strftime("%Y-%m")

    from collections import Counter
    rank_posts  = Counter()
    rank_users  = {}
    total_posts = 0

    boards = [board_id] if board_id else KR_BOARDS_FOR_RANK

    def get_group(rank_name):
        rn = rank_name.lower()
        if "active" in rn:             return "Active"
        elif "beginner" in rn:         return "Beginner"
        elif "expert" in rn:           return "Expert"
        elif "★" in rank_name:         return "Stars"
        elif "manager" in rn or "moderator" in rn: return "Staff"
        else:                          return "기타"

    for bid in boards:
        offset = 0
        while offset <= max_scan:
            q = (f"SELECT id, post_time, author.id, author.rank "
                 f"FROM messages WHERE board.id = '{bid}' AND depth = 0 "
                 f"ORDER BY post_time DESC LIMIT 50 OFFSET {offset}")
            data  = liql(q)
            items = data.get("items", [])
            if not items: break

            stop = False
            for item in items:
                pt = item.get("post_time","")[:7]
                if pt > dt_to.strftime("%Y-%m"):
                    continue
                if pt < dt_from.strftime("%Y-%m"):
                    stop = True
                    break

                dt = parse_dt(item.get("post_time",""))
                if dt is None: continue
                dt_utc = dt.astimezone(timezone.utc)
                if dt_utc < dt_from or dt_utc > dt_to:
                    continue

                author    = item.get("author", {}) or {}
                rank_info = author.get("rank", {}) or {}
                rank_name = rank_info.get("name", "Unknown")
                user_id   = author.get("id", "")
                group     = get_group(rank_name)

                rank_posts[group] += 1
                if group not in rank_users:
                    rank_users[group] = set()
                rank_users[group].add(user_id)
                total_posts += 1

            if stop: break
            offset += 50
            time.sleep(0.1)

    results = sorted([
        {
            "rank":     rank,
            "posts":    cnt,
            "users":    len(rank_users.get(rank, set())),
            "user_ids": list(rank_users.get(rank, set())),
            "avg":      round(cnt / len(rank_users.get(rank, set())), 1) if rank_users.get(rank) else 0,
        }
        for rank, cnt in rank_posts.items()
    ], key=lambda x: -x["posts"])

    return jsonify({
        "period":      {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        "total_posts": total_posts,
        "total_users": len(set(uid for s in rank_users.values() for uid in s)),
        "results":     results,
        "boards":      boards,
    })


# ── 커뮤니티 키워드 검색 ─────────────────────────────────────────────────────────
@app.route("/api/community/search", methods=["POST"])
def community_keyword_search():
    body     = request.json or {}
    keyword  = body.get("keyword", "").strip()
    date_from = body.get("from", "")
    date_to   = body.get("to", "")
    kr_only   = body.get("kr_only", True)
    max_results = int(body.get("max_results", 200))

    if not keyword:
        return jsonify({"error": "키워드를 입력해주세요"}), 400

    try:
        if date_from:
            dt_from = datetime.fromisoformat(date_from).replace(tzinfo=timezone.utc)
        else:
            dt_from = datetime.now(timezone.utc) - timedelta(days=30)
        if date_to:
            dt_to = datetime.fromisoformat(date_to).replace(hour=23, minute=59, second=59, tzinfo=timezone.utc)
        else:
            dt_to = datetime.now(timezone.utc)
    except:
        return jsonify({"error": "날짜 형식 오류"}), 400

    results   = []
    offset    = 0
    limit     = 50
    board_map = {
        "kr-community-sp-gs":    "갤럭시 S",
        "kr-community-sp-gfold": "갤럭시 Z",
        "kr-community-wear":     "웨어러블",
        "kr-community-sp-ga":    "갤럭시 A",
        "kr-community-svc-other":"커뮤니티",
        "kr-community-other":    "커뮤니티",
        "kr-community-promo":    "프로모션",
        "kr-community-pc":       "PC/태블릿",
        "kr-community-home":     "홈가전",
        "kr-news":               "뉴스",
        "kr-community-sp-gtab":  "태블릿",
        "kr-contest-curator250702": "큐레이터",
    }

    while len(results) < max_results:
        # 띄어쓰기 있으면 구문 검색("..."), 없으면 단어 검색
        if ' ' in keyword:
            search_term = f'"{keyword}"'
        else:
            search_term = keyword
        q = (f"SELECT id, subject, post_time, view_href, board FROM messages "
             f"WHERE depth = 0 AND subject MATCHES '{search_term}' "
             f"ORDER BY post_time DESC LIMIT {limit} OFFSET {offset}")
        data  = liql(q)
        items = data.get("items", [])
        if not items:
            break

        all_before_range = True  # 이 배치가 전부 기간 이전인지 체크

        for item in items:
            try:
                board_raw = item.get("board", {})
                board_id  = board_raw.get("id", "") if isinstance(board_raw, dict) else ""
                if kr_only and not board_id.startswith("kr-"):
                    continue

                dt = parse_dt(item["post_time"])
                if dt is None:
                    continue
                dt_utc = dt.astimezone(timezone.utc)

                # 기간 이후 → 스킵 (최신순이므로 초반에만 해당)
                if dt_utc > dt_to:
                    all_before_range = False
                    continue

                # 기간 이전 → 더 오래된 건 없으므로 중단 플래그
                if dt_utc < dt_from:
                    continue  # 개별 항목은 스킵, 배치 전체 판단은 아래서

                all_before_range = False  # 기간 내 항목 존재
                board_name = board_map.get(board_id, board_id or "기타")
                results.append({
                    "id":       item["id"],
                    "subject":  item.get("subject", ""),
                    "date":     dt_utc.strftime("%Y-%m-%d"),
                    "time":     dt_utc.strftime("%H:%M"),
                    "board":    board_name,
                    "board_id": board_id,
                    "url":      item.get("view_href", ""),
                })
            except:
                continue

        # 이 배치의 가장 오래된 항목이 기간 시작보다 이전이면 중단
        try:
            oldest_in_batch = parse_dt(items[-1]["post_time"])
            if oldest_in_batch:
                oldest_utc = oldest_in_batch.astimezone(timezone.utc)
                if oldest_utc < dt_from:
                    break
        except:
            pass

        if len(items) < limit:
            break
        offset += limit
        time.sleep(0.15)

    results.sort(key=lambda x: x["date"] + x["time"], reverse=True)

    # 게시판별 집계
    board_count = {}
    for r in results:
        b = r["board"]
        board_count[b] = board_count.get(b, 0) + 1
    board_count = dict(sorted(board_count.items(), key=lambda x: x[1], reverse=True))

    return jsonify({
        "keyword":     keyword,
        "period":      {"from": dt_from.strftime("%Y-%m-%d"), "to": dt_to.strftime("%Y-%m-%d")},
        "kr_only":     kr_only,
        "total":       len(results),
        "board_count": board_count,
        "posts":       results,
    })


# ── Instagram 스토리 수집 ─────────────────────────────────────────────────────
@app.route("/api/apify/instagram/stories", methods=["POST"])
def get_instagram_stories():
    body     = request.json or {}
    username = body.get("username","").strip()
    if not username:
        return jsonify({"error": "username 필요"}), 400

    try:
        url = (f"{APIFY_BASE}/acts/apify~instagram-scraper"
               f"/run-sync-get-dataset-items?token={APIFY_API_KEY}&timeout=120")
        resp = requests.post(url,
            headers={"Content-Type": "application/json"},
            json={
                "directUrls":   [f"https://www.instagram.com/{username}/"],
                "resultsType":  "stories",
                "resultsLimit": 50,
                "searchType":   "hashtag",
                "addParentData": False,
            }, timeout=150)

        if resp.status_code not in (200, 201):
            return jsonify({"error": f"Apify 오류: {resp.status_code}"}), 500

        items = resp.json()
        if not isinstance(items, list):
            return jsonify({"error": "결과 조회 실패"}), 500

        stories = []
        for p in items:
            try:
                # 오류 항목 스킵
                if p.get("error") or not p.get("timestamp"):
                    continue
                ts  = p.get("timestamp","")
                dt  = datetime.fromisoformat(ts.replace("Z","+00:00")).astimezone(timezone.utc) if ts else None
                stories.append({
                    "id":        p.get("id",""),
                    "date":      dt.strftime("%Y-%m-%d") if dt else "",
                    "time":      dt.strftime("%H:%M") if dt else "",
                    "type":      p.get("type",""),
                    "duration":  round(p.get("videoDuration",0) or 0, 1),
                    "thumbnail": p.get("displayUrl",""),
                    "url":       p.get("url","") or f"https://www.instagram.com/{username}/",
                    "is_highlight": bool(p.get("isPinned")),
                })
            except: continue

        stories.sort(key=lambda x: x["date"]+x["time"], reverse=True)

        return jsonify({
            "username": username,
            "total":    len(stories),
            "stories":  stories,
        })
    except requests.Timeout:
        return jsonify({"error": "수집 타임아웃"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── 실시간 모니터링 API ───────────────────────────────────────────────────────
@app.route("/api/monitor/community", methods=["POST"])
def monitor_community():
    """스타즈 전체 커뮤니티 최신 게시물 수집 (최근 N시간)"""
    body    = request.json or {}
    user_ids = body.get("user_ids", [])
    hours   = int(body.get("hours", 24))

    dt_from = datetime.now(timezone.utc) - timedelta(hours=hours)

    results = {}
    for uid in user_ids:
        posts = []
        q = (f"SELECT id, subject, post_time, view_href, board FROM messages "
             f"WHERE author.id = '{uid}' AND depth = 0 "
             f"ORDER BY post_time DESC LIMIT 20")
        data  = liql(q)
        items = data.get("items", [])

        for item in items:
            dt = parse_dt(item["post_time"])
            if dt is None: continue
            dt_utc = dt.astimezone(timezone.utc)
            if dt_utc < dt_from: break

            board_raw = item.get("board", {})
            board_id  = board_raw.get("id", "") if isinstance(board_raw, dict) else ""
            board_map = {
                "kr-community-sp-gs":    "갤럭시 S",
                "kr-community-sp-gfold": "갤럭시 Z",
                "kr-community-wear":     "웨어러블",
                "kr-community-sp-ga":    "갤럭시 A",
                "kr-community-svc-other":"커뮤니티",
                "kr-community-other":    "커뮤니티",
                "kr-community-promo":    "프로모션",
                "kr-news":               "뉴스",
            }
            posts.append({
                "id":      item["id"],
                "subject": item.get("subject",""),
                "date":    dt_utc.strftime("%m/%d %H:%M"),
                "board":   board_map.get(board_id, board_id or "기타"),
                "url":     item.get("view_href",""),
            })
        results[uid] = posts
        time.sleep(0.2)

    return jsonify({
        "updated_at": datetime.now(timezone.utc).strftime("%H:%M:%S"),
        "hours":      hours,
        "results":    results,
    })


@app.route("/api/monitor/instagram/single", methods=["POST"])
def monitor_instagram_single():
    """단일 계정 Instagram 게시물 + 스토리 수집"""
    body     = request.json or {}
    username = body.get("username","").strip()
    login    = body.get("login","")
    hours    = int(body.get("hours", 24))

    if not username:
        return jsonify({"error": "username 필요"}), 400

    dt_from = datetime.now(timezone.utc) - timedelta(hours=hours)
    posts   = []
    stories = []

    sync_url = (f"{APIFY_BASE}/acts/apify~instagram-scraper"
                f"/run-sync-get-dataset-items?token={APIFY_API_KEY}&timeout=120")

    # 1. 게시물 수집
    try:
        resp = requests.post(sync_url,
            headers={"Content-Type": "application/json"},
            json={
                "directUrls":   [f"https://www.instagram.com/{username}/"],
                "resultsType":  "posts",
                "resultsLimit": 10,
                "searchType":   "hashtag",
                "addParentData": False,
            }, timeout=150)
        items = resp.json() if resp.status_code in (200, 201) else []
        for p in (items if isinstance(items, list) else []):
            try:
                dt = datetime.fromisoformat(p["timestamp"].replace("Z","+00:00")).astimezone(timezone.utc)
                if dt < dt_from: continue
                raw_type     = p.get("type","")
                product_type = p.get("productType","") or ""
                if product_type.lower() == "clips": post_type = "Reels"
                elif raw_type == "Video": post_type = "동영상"
                elif raw_type == "Sidecar": post_type = "슬라이드"
                else: post_type = "이미지"
                posts.append({
                    "date":     dt.strftime("%m/%d %H:%M"),
                    "type":     post_type,
                    "caption":  (p.get("caption","") or "")[:60],
                    "likes":    p.get("likesCount",0) or 0,
                    "comments": p.get("commentsCount",0) or 0,
                    "views":    p.get("videoViewCount",0) or 0,
                    "url":      p.get("url",""),
                    "thumbnail": p.get("displayUrl",""),
                })
            except: continue
    except: pass

    time.sleep(1)

    # 2. 스토리 수집
    try:
        resp2 = requests.post(sync_url,
            headers={"Content-Type": "application/json"},
            json={
                "directUrls":   [f"https://www.instagram.com/{username}/"],
                "resultsType":  "stories",
                "resultsLimit": 20,
                "searchType":   "hashtag",
                "addParentData": False,
            }, timeout=120)
        story_items = resp2.json() if resp2.status_code in (200, 201) else []
        for p in (story_items if isinstance(story_items, list) else []):
            try:
                dt = datetime.fromisoformat(p["timestamp"].replace("Z","+00:00")).astimezone(timezone.utc)
                if dt < dt_from: continue
                stories.append({
                    "date":      dt.strftime("%m/%d %H:%M"),
                    "type":      p.get("type",""),
                    "duration":  p.get("videoDuration",0) or 0,
                    "thumbnail": p.get("displayUrl",""),
                    "url":       p.get("url","") or f"https://www.instagram.com/{username}/",
                })
            except: continue
    except: pass

    return jsonify({
        "login":      login,
        "username":   username,
        "posts":      posts,
        "stories":    stories,
        "updated_at": datetime.now(timezone.utc).strftime("%H:%M:%S"),
    })


@app.route("/api/monitor/instagram", methods=["POST"])
def monitor_instagram():
    """스타즈 Instagram 최신 게시물 수집 (다중 - 순차)"""
    body     = request.json or {}
    accounts = body.get("accounts", [])
    hours    = int(body.get("hours", 24))
    results  = {}

    for acc in accounts:
        username = acc.get("username","").strip()
        login    = acc.get("login","")
        if not username:
            continue
        try:
            r = requests.post(
                f"http://localhost:{os.environ.get('PORT',5000)}/api/monitor/instagram/single",
                json={"username": username, "login": login, "hours": hours},
                headers={"Content-Type": "application/json"},
                timeout=300
            )
            data = r.json()
            results[login] = {
                "posts":   data.get("posts",[]),
                "stories": data.get("stories",[]),
            }
        except Exception as e:
            results[login] = {"posts": [], "stories": []}
        time.sleep(0.5)

    return jsonify({
        "updated_at": datetime.now(timezone.utc).strftime("%H:%M:%S"),
        "hours":      hours,
        "results":    results,
    })


# ── 정적 파일 ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")

if __name__ == "__main__":
    os.makedirs("static", exist_ok=True)
    print("\n🚀 Samsung Members 대시보드 서버 시작")
    print("📌 브라우저에서 접속: http://localhost:5000\n")
    app.run(host="0.0.0.0", debug=False, port=5000)
