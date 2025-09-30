#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KeepAlive_Combined_Windows.py
Phiên bản Windows: fix rclone, giữ Admin tasks, thêm gửi mail Copilot ra ngoài.
"""

import os, sys, json, random, argparse, requests, feedparser, logging

# --- Fix UnicodeEncodeError on Windows logging ---
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from flask import Flask, redirect, request

# ------------------ Logging ------------------
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Ép logger dùng UTF-8, nếu ký tự không encode được thì thay bằng '?'
try:
    logging.getLogger().handlers[0].stream.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
# -------------------------------------------------

def log(msg, level="info"):
    getattr(logging, level)(msg)

# ------------------ Load .env ------------------
load_dotenv()
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL")
USER_EMAIL = os.getenv("USER_EMAIL")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:8000/callback")
IMAGE_FOLDER = os.getenv("IMAGE_FOLDER", "C:\office365E5\images")
RCLONE_REMOTE = os.getenv("RCLONE_REMOTE", "onedrive")
RCLONE_CLEAN_FOLDER = os.getenv("RCLONE_CLEAN_FOLDER", "KeepAliveClean")
LOCAL_UPLOAD = os.getenv("LOCAL_UPLOAD", "upload_local")
REMOTE_UPLOAD = os.getenv("REMOTE_UPLOAD", "backup_test")

# Danh sách mail ngoài (cách nhau bởi dấu phẩy)
EXTERNAL_EMAILS = os.getenv("EXTERNAL_EMAILS", "")
EXTERNAL_EMAILS = [e.strip() for e in EXTERNAL_EMAILS.split(",") if e.strip()]

TOKEN_FILE = "token.json"

# ------------------ Token helpers ------------------
def save_token(token_data):
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(token_data, f, indent=2)
    log("🔑 Lưu token.json thành công")

def load_token():
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None

def refresh_access_token(refresh_token):
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
        "refresh_token": refresh_token, "grant_type": "refresh_token",
        "redirect_uri": REDIRECT_URI
    }
    res = requests.post(url, data=data)
    if res.status_code == 200:
        token_json = res.json()
        save_token(token_json)
        log("🔄 Refresh token thành công")
        return token_json.get("access_token")
    else:
        log(f"❌ Refresh token lỗi: {res.text}", "error")
        return None

# ------------------ Token App/User ------------------
def get_token_app():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {"client_id": CLIENT_ID,"client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default","grant_type": "client_credentials"}
    res = requests.post(url, data=data)
    if res.status_code != 200:
        log("❌ Không lấy được token App", "error"); sys.exit(1)
    return res.json()["access_token"]

def get_token_user_flask():
    app = Flask(__name__)
    scopes = [
        "offline_access",
        "https://graph.microsoft.com/Mail.Send",
        "https://graph.microsoft.com/User.Read",
        "https://graph.microsoft.com/Files.ReadWrite",
        "https://graph.microsoft.com/Calendars.ReadWrite"
    ]
    authorize_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    @app.route("/")
    def home():
        return redirect(f"{authorize_url}?client_id={CLIENT_ID}&response_type=code&redirect_uri={REDIRECT_URI}&scope={' '.join(scopes)}")

    @app.route("/callback")
    def callback():
        code = request.args.get("code")
        token_res = requests.post(token_url, data={
            "client_id": CLIENT_ID, "scope": " ".join(scopes),
            "code": code, "redirect_uri": REDIRECT_URI,
            "grant_type": "authorization_code", "client_secret": CLIENT_SECRET
        })
        token_json = token_res.json()
        if "access_token" not in token_json: return "❌ Lỗi lấy token"
        save_token(token_json)
        run_tasks(token_json["access_token"], user_mode=True)
        return "✅ Đăng nhập & chạy xong, xem log tại ping_log.txt"

    log("⚡ Flask login tại http://localhost:8000 ...")
    app.run(host="0.0.0.0", port=8000)

def get_token_user():
    token_data = load_token()
    if token_data and "refresh_token" in token_data:
        return refresh_access_token(token_data["refresh_token"])
    else:
        get_token_user_flask()
        return None

# ------------------ Graph API helpers ------------------
def get_users(token):
    r = requests.get("https://graph.microsoft.com/v1.0/users", headers={"Authorization": f"Bearer {token}"})
    if r.status_code == 200:
        users = r.json().get("value", [])
        log(f"✅ Lấy danh sách user: {len(users)}")
        return users
    else:
        log(f"❌ Lấy users lỗi: {r.status_code} {r.text}", "error")
        return []

# ------------------ Basic Tasks ------------------
def send_ping_mail(token, user_mode=False):
    if not USER_EMAIL:
        return log("⚠️ USER_EMAIL chưa cấu hình", "warning")
    mail_url = "https://graph.microsoft.com/v1.0/me/sendMail" if user_mode \
             else f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/sendMail"
    payload = {"message": {"subject": "Ping Mail giữ tài khoản sống",
              "body": {"contentType": "Text", "content": "Test gửi mail tự động"},
              "toRecipients": [{"emailAddress": {"address": USER_EMAIL}}]}}
    r = requests.post(mail_url, headers={"Authorization": f"Bearer {token}"}, json=payload)
    log(f"📧 Mail gửi: {r.status_code}")

def upload_pingalive(token, user_mode=False):
    if not USER_EMAIL: return log("⚠️ USER_EMAIL chưa cấu hình", "warning")
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/PingAlive.txt:/content" if user_mode \
        else f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}/drive/root:/PingAlive.txt:/content"
    r = requests.put(url, headers={"Authorization": f"Bearer {token}"}, data="KeepAlive".encode())
    log(f"📄 PingAlive.txt: {r.status_code}")

def rclone_tasks(skip=False):
    if skip: return log("⏭️ Bỏ qua rclone (--skip-rclone)")
    log(f"🧹 Xoá nội dung {RCLONE_REMOTE}:{RCLONE_CLEAN_FOLDER}")
    os.system(f"rclone delete {RCLONE_REMOTE}:{RCLONE_CLEAN_FOLDER}")
    if os.path.exists(LOCAL_UPLOAD):
        log(f"📂 Upload {LOCAL_UPLOAD}/ → {RCLONE_REMOTE}:{REMOTE_UPLOAD}")
        os.system(f'rclone copy "{LOCAL_UPLOAD}" {RCLONE_REMOTE}:{REMOTE_UPLOAD} --transfers=4 --checkers=8 --fast-list')
    else:
        log(f"⚠️ Không có thư mục {LOCAL_UPLOAD}", "warning")

# ------------------ Advanced Tasks ------------------
def create_daily_event(token, user_id):
    now = datetime.now(timezone.utc)
    start = now.replace(hour=9, minute=0, second=0, microsecond=0)
    end = start + timedelta(minutes=30)
    payload = {"subject": "📌 Daily Auto Event",
        "start": {"dateTime": start.isoformat(), "timeZone": "UTC"},
        "end": {"dateTime": end.isoformat(), "timeZone": "UTC"}}
    r = requests.post(f"https://graph.microsoft.com/v1.0/users/{user_id}/events",
                      headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"}, json=payload)
    log(f"📆 Tạo sự kiện: {r.status_code}")

def get_news_rss(limit=5):
    feed = feedparser.parse("https://vnexpress.net/rss/tin-moi-nhat.rss")
    return "\n".join(f"- {e.title}" for e in feed.entries[:limit])

def generate_copilot_mock():
    return random.choice([
        "🧠 Hôm nay trời nắng.",
        "📈 VN-Index tăng nhẹ.",
        "💡 Mẹo: Ctrl+Shift+V để dán không định dạng."
    ])

def send_personalized_mails(token, sender_email, recipients, user_id):
    # gộp recipients nội bộ và mail ngoài
    all_recipients = recipients + EXTERNAL_EMAILS
    for rcp in all_recipients:
        body = f"📢 Tin mới:\n{get_news_rss()}\n\n🤖 Copilot:\n{generate_copilot_mock()}"
        payload = {"message": {"subject": "📌 Bản tin & Copilot",
                   "body": {"contentType": "Text", "content": body},
                   "toRecipients": [{"emailAddress": {"address": rcp}}]}}
        res = requests.post(f"https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail",
                            headers={"Authorization": f"Bearer {token}"}, json=payload)
        log(f"📧 Gửi {rcp}: {res.status_code}")

def upload_random_images(token, user_id, folder="E5Auto"):
    if not os.path.exists(IMAGE_FOLDER): return log("⚠️ Không có thư mục ảnh", "warning")
    imgs = [f for f in os.listdir(IMAGE_FOLDER) if f.lower().endswith(('.jpg','.png'))]
    if not imgs: return log("⚠️ Không có ảnh", "warning")
    for f in random.sample(imgs, min(3,len(imgs))):
        with open(os.path.join(IMAGE_FOLDER,f),"rb") as fd:
            url=f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/{folder}/{f}:/content"
            r=requests.put(url,headers={"Authorization":f"Bearer {token}"},data=fd)
        log(f"🖼️ Upload {f}: {r.status_code}")

# ------------------ Run tasks ------------------
def run_tasks(token, user_mode=False, skip_rclone=False):
    send_ping_mail(token, user_mode)
    upload_pingalive(token, user_mode)
    rclone_tasks(skip_rclone)

    if not user_mode:
        users = get_users(token)
        admin = next((u for u in users if u["userPrincipalName"].lower()==ADMIN_EMAIL.lower()), None)
        if admin:
            aid = admin["id"]
            create_daily_event(token, aid)
            recps = [u["userPrincipalName"] for u in users if u["userPrincipalName"]!=ADMIN_EMAIL]
            send_personalized_mails(token, ADMIN_EMAIL, recps, aid)
            upload_random_images(token, aid)
    log("✅ Hoàn tất tất cả tác vụ")

# ------------------ Main ------------------
if __name__=="__main__":
    p=argparse.ArgumentParser()
    p.add_argument("--app",action="store_true"); p.add_argument("--user",action="store_true")
    p.add_argument("--skip-rclone",action="store_true")
    a=p.parse_args()
    if a.app: run_tasks(get_token_app(), user_mode=False, skip_rclone=a.skip_rclone)
    elif a.user:
        t=get_token_user()
        if t: run_tasks(t, user_mode=True, skip_rclone=a.skip_rclone)
    else: log("❌ Chọn --app hoặc --user","error")
