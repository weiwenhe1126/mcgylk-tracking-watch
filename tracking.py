"""铭川跨境物流跟单查询（邮件触发，非定时推送）。

机制：
  每 5 分钟检查一次收件箱，发现「本邮箱自己发出的、标题不含【铭川跟单】」的新邮件，
  从中提取 10~20 位数字作为单号，查询后直接回信给发件人。

  回信邮件带 X-Tracking-Bot 头且标题以【铭川跟单】开头，会被自动跳过，不会死循环。

用法：
  python tracking.py            轮询并回复
  python tracking.py --dry-run  只打印，不写状态、不发信
"""

import email
import base64
import hashlib
import hmac
import imaplib
import json
import os
import re
import smtplib
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.mime.text import MIMEText

API = "http://mcgylk.nextsls.com/tracking/app?inajax=1&tracking_number={}"
STATE = ".tracking_state.json"
MAX_NUMBERS = 20
MAX_MAILS_PER_RUN = 10
BOOTSTRAP_LOOKBACK = 20

STATUS_MAP = {
    "in_transit": "转运中",
    "delivered": "已签收",
    "out_for_delivery": "派送中",
    "pending": "待处理",
    "created": "已下单",
    "exception": "异常",
    "returned": "退回",
    "customs": "清关",
    "arrived": "已到港",
}


def decode_mime(value):
    if not value:
        return ""
    out = []
    for part, enc in decode_header(value):
        if isinstance(part, bytes):
            try:
                out.append(part.decode(enc or "utf-8", "ignore"))
            except Exception:
                out.append(part.decode("utf-8", "ignore"))
        else:
            out.append(part)
    return "".join(out)


def message_text(msg):
    texts = []
    for part in msg.walk():
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        try:
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
            if payload:
                texts.append(payload.decode(charset, "ignore"))
        except Exception:
            continue
    return "\n".join(texts)


def extract_numbers(text):
    return set(re.findall(r"(?<!\d)\d{10,20}(?!\d)", text or ""))


def collect_numbers(user, auth, last_uid):
    """读取新邮件，返回 (单号集合, 本次见到的最新 UID, 命中的邮件数)。"""
    numbers = set()
    hits = 0
    self_key = user.split("@")[0].lower()
    bootstrap = last_uid == 0

    box = imaplib.IMAP4_SSL("imap.qq.com", 993)
    try:
        box.login(user, auth)
        box.select("INBOX")
        typ, data = box.uid("search", None, "ALL")
        if typ != "OK" or not data or not data[0]:
            print("INBOX_EMPTY")
            return numbers, last_uid, hits
        all_uids = sorted(int(x) for x in data[0].split())
        max_uid = all_uids[-1]

        if bootstrap:
            # 首次运行只看最近 20 封，避免把历史邮件全部回一遍
            batch = all_uids[-BOOTSTRAP_LOOKBACK:]
        else:
            batch = [u for u in all_uids if u > last_uid]
        batch = batch[-MAX_MAILS_PER_RUN:]

        for uid in batch:
            typ, raw = box.uid("fetch", str(uid), "(RFC822)")
            if typ != "OK" or not raw or not isinstance(raw[0], tuple):
                continue
            msg = email.message_from_bytes(raw[0][1])
            frm = (msg.get("From") or "").lower()
            if self_key not in frm:
                continue
            subject = decode_mime(msg.get("Subject"))
            # 跳过自己的回信，避免循环
            if subject.startswith("\u3010\u94ed\u5ddd\u8ddf\u5355\u3011") or msg.get("X-Tracking-Bot"):
                continue
            found = extract_numbers(subject + "\n" + message_text(msg))
            if found:
                hits += 1
                print(f"MAIL uid={uid} subject={subject!r} numbers={sorted(found)}")
                numbers |= found
    finally:
        try:
            box.logout()
        except Exception:
            pass
    return numbers, max_uid, hits


def query(number):
    req = urllib.request.Request(
        API.format(number), headers={"User-Agent": "Mozilla/5.0"}
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        payload = json.loads(resp.read().decode("utf-8", "ignore"))
    if payload.get("status") != 1 or not payload.get("data"):
        return None
    return payload["data"].get("shipment")


def load_state():
    if os.path.exists(STATE):
        try:
            with open(STATE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_state(state):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


def build_report(items, now_str):
    lines = [f"铭川跨境物流跟单查询结果（{now_str} 北京时间）", ""]
    for number, ship in items:
        if not ship:
            lines += [f"单号：{number}", "状态：查询不到该单号（可能单号有误或还未入系统）", ""]
            continue
        st = ship.get("status", "")
        lines.append(f"单号：{number}")
        lines.append(f"状态：{STATUS_MAP.get(st, st)}（{st}）")
        lines.append(
            f"目的国：{ship.get('country', '-')}  邮编：{ship.get('postcode', '-')}  "
            f"包裹数：{ship.get('parcel_count', '-')}"
        )
        if ship.get("ext_number"):
            lines.append(f"参考号：{ship.get('ext_number')}")
        traces = ship.get("traces") or []
        lines.append("最新轨迹：")
        for t in traces[:5]:
            lines.append(f"  · {t.get('time', '')}  {t.get('info', '')}")
        lines.append("")
    lines.append("查询入口：http://mcgylk.nextsls.com/tracking/app#/tracking")
    lines.append("")
    lines.append("提示：把单号写进邮件发到本邮箱即可，系统每 5 分钟检查一次并自动回信。")
    return "\n".join(lines)


def send_mail(subject, body):
    user = os.environ.get("MAIL_USER", "").strip()
    auth = os.environ.get("MAIL_AUTH", "").strip()
    to = os.environ.get("MAIL_TO", "").strip() or user
    if not user or not auth:
        raise RuntimeError("缺少 MAIL_USER / MAIL_AUTH 环境变量")
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = decode_mime(subject)
    msg["From"] = user
    msg["To"] = to
    msg["X-Tracking-Bot"] = "1"
    with smtplib.SMTP_SSL("smtp.qq.com", 465, timeout=60) as srv:
        srv.login(user, auth)
        srv.sendmail(user, [x.strip() for x in to.split(",") if x.strip()], msg.as_string())
    print(f"MAIL_SENT -> {to}")


def ding_markdown(body):
    """邮件正文 -> 钉钉 markdown：内容与邮件一致，仅按钉钉要求改用空行分段。"""
    return "\n\n".join(line.strip() for line in body.splitlines() if line.strip())


def notify_ding(title, body):
    """钉钉群机器人推送（加签），正文与邮件保持一致。未配置或失败时只打印，不影响回信。"""
    hook = os.environ.get("DING_WEBHOOK", "").strip()
    secret = os.environ.get("DING_SECRET", "").strip()
    if not hook:
        print("DING_SKIP: 未配置 DING_WEBHOOK")
        return
    try:
        ts = str(int(time.time() * 1000))
        digest = hmac.new(
            secret.encode(), f"{ts}\n{secret}".encode(), hashlib.sha256
        ).digest()
        sign = urllib.parse.quote_plus(base64.b64encode(digest).decode())
        text = f"### {title}\n\n" + ding_markdown(body)
        payload = json.dumps(
            {"msgtype": "markdown", "markdown": {"title": title, "text": text}}
        ).encode()
        req = urllib.request.Request(
            f"{hook}&timestamp={ts}&sign={sign}",
            data=payload,
            headers={"Content-Type": "application/json"},
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode("utf-8", "ignore"))
        print(f"DING_SENT errcode={result.get('errcode')} errmsg={result.get('errmsg')}")
    except Exception as exc:
        print(f"DING_ERROR {exc}")


def main():
    dry = "--dry-run" in sys.argv
    user = os.environ.get("MAIL_USER", "").strip()
    auth = os.environ.get("MAIL_AUTH", "").strip()
    if not user or not auth:
        raise RuntimeError("缺少 MAIL_USER / MAIL_AUTH 环境变量")

    state = load_state()
    last_uid = int(state.get("last_uid") or 0)
    print(f"last_uid={last_uid}")

    numbers, max_uid, hits = collect_numbers(user, auth, last_uid)
    print(f"mail_hits={hits} numbers={sorted(numbers)}")

    if not numbers:
        print("没有新的单号邮件，退出")
        if not dry and max_uid > last_uid:
            save_state({"last_uid": max_uid})
        return

    numbers = sorted(numbers)[:MAX_NUMBERS]
    items = []
    for n in numbers:
        try:
            ship = query(n)
        except Exception as exc:
            print(f"QUERY_ERROR {n}: {exc}")
            ship = None
        items.append((n, ship))

    now_str = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")
    body = build_report(items, now_str)

    if len(items) == 1:
        n, ship = items[0]
        st = ship.get("status", "") if ship else ""
        subject = f"【铭川跟单】{n} {STATUS_MAP.get(st, '查询不到')}"
    else:
        subject = f"【铭川跟单】{len(items)} 票查询结果（{now_str[5:]}）"

    print(subject)
    print(body)
    if dry:
        print("DRY_RUN: 未发信")
        return
    send_mail(subject, body)
    notify_ding(subject, body)
    save_state({"last_uid": max_uid})


if __name__ == "__main__":
    main()
