"""
Teams Web会議マニュアル案内チャットを M365 全員に送信。

【新規ユーザー追加時の再利用方法】
  ・EXCLUDE_UPN に admin と除外したい人を追加
  ・SINGLE_TARGET_UPN を指定すれば1名だけに送信できる（新規入社者用）
  ・SINGLE_TARGET_UPN=None なら EXCLUDE_UPN 以外の全員に送信

送信元: h.hasebe@besterra.onmicrosoft.com（長谷部本人）
方式: Microsoft Graph PowerShell clientId で ROPC、テナント admin consent 済
本文: 2026-05-11 rev3 文面（5セクション構成のマニュアル案内）
"""
from __future__ import annotations

import html as h
import json
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

import msal

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).parent
MAPPING = ROOT.parent / "m365_admin_access_20260501" / "m365_smarthr_mapping.json"
# 上記が無い時のフォールバック（ClaudeCode 配下にもある）
FALLBACK_MAPPING = Path(r"C:\Users\h.hasebe\ClaudeCode\m365_admin_access_20260501\m365_smarthr_mapping.json")
LOG = ROOT / "send_webmeeting_log.json"

TENANT = "besterra.onmicrosoft.com"
SELF_UPN = "h.hasebe@besterra.onmicrosoft.com"
SELF_PW = "8esterrA99"
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
SCOPES = [
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/User.ReadBasic.All",
]
GRAPH = "https://graph.microsoft.com/v1.0"

# ----- 配信制御 -----
EXCLUDE_UPN = {
    "admin@besterra.onmicrosoft.com",
    "k.kanemitsu@besterra.onmicrosoft.com",  # 金光さん
}
# 新規ユーザー追加時はここに UPN を入れると1名だけ送信
SINGLE_TARGET_UPN = None  # 例: "t.watanabe@besterra.onmicrosoft.com"

MANUAL_URL = "https://hhasebe-besterra.github.io/teams-manual/"

BODY = f"""Teams の Web 会議（オンラインミーティング）の使い方マニュアルを公開しました。

これまでベステラの Web 会議は Zoom の有料アカウントを共有する形でしたが、皆さんに M365 Business Standard ライセンスをお配りしたので、これからは Teams から各自のアカウントで自由に会議を開けます。

ただ Teams での Web 会議に慣れていない方が多いと思いますので、画面イメージ付きで「開く」「招く」「入る」「便利機能」「会社携帯への導入」の5ステップにまとめたガイドを作りました。

▼ ベステラ Teams Web会議マニュアル
{MANUAL_URL}

社外の方を招待する手順、画面共有、録画はもちろん、「一人で会議を開いて自分の声を録画する」（議事メモや備忘録に便利）方法も載せています。会社携帯（iPhone / Android）への Teams アプリ導入手順もあるので、外出先や移動中でも会議に参加できます。

なお、社内のメインのコミュニケーション基盤は当面 Slack のままで、Teams は段階的に併用 → 移行していく形になります。今すぐ全部を Teams に切り替えるという話ではなく、Web 会議から少しずつ慣れていただければと思います。

ご不明な点はこのチャットに返信してください。"""


def to_html(plain: str) -> str:
    out = []
    for line in plain.splitlines():
        e = h.escape(line)
        if e.startswith("▼"):
            e = f"<b>{e}</b>"
        # URL を a タグに
        if "https://" in e:
            i = e.find("https://")
            url = e[i:].split()[0]
            e = e[:i] + f'<a href="{url}">{url}</a>'
        out.append(e)
    return "<br>".join(out)


def graph(t, method, path, body=None):
    url = path if path.startswith("http") else GRAPH + path
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, headers={
        "Authorization": f"Bearer {t}", "Content-Type": "application/json"
    }, method=method)
    try:
        with urllib.request.urlopen(req) as r:
            raw = r.read()
            return json.loads(raw) if raw else None
    except urllib.error.HTTPError as e:
        body_txt = e.read().decode("utf-8", errors="replace")
        e.body = body_txt
        raise


def main():
    mapping_path = MAPPING if MAPPING.exists() else FALLBACK_MAPPING
    with mapping_path.open(encoding="utf-8") as f:
        mp = json.load(f)

    if SINGLE_TARGET_UPN:
        targets = [m for m in mp["mapped"] if m["m365_upn"] == SINGLE_TARGET_UPN]
    else:
        targets = [m for m in mp["mapped"] if m["m365_upn"] not in EXCLUDE_UPN]

    print(f"recipients: {len(targets)}")
    for t in targets:
        print(f"  - {t.get('sh_name') or t.get('m365_displayName')} <{t['m365_upn']}>")

    app = msal.PublicClientApplication(CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT}")
    tk = app.acquire_token_by_username_password(SELF_UPN, SELF_PW, scopes=SCOPES)
    if "access_token" not in tk:
        print("AUTH FAILED:", json.dumps(tk, ensure_ascii=False, indent=2))
        sys.exit(1)
    t = tk["access_token"]

    me = graph(t, "GET", "/me?$select=id,userPrincipalName,displayName")
    print(f"\nsender: {me['displayName']} <{me['userPrincipalName']}>")

    html_body = to_html(BODY)

    log = []
    for i, r in enumerate(targets, 1):
        upn = r["m365_upn"]
        if upn == SELF_UPN:
            print(f"  [{i:2d}/{len(targets)}] SKIP self")
            log.append({"upn": upn, "skip_self": True})
            continue
        try:
            tgt = graph(t, "GET", f"/users/{upn}?$select=id,userPrincipalName,displayName")
            chat = graph(t, "POST", "/chats", {
                "chatType": "oneOnOne",
                "members": [
                    {"@odata.type": "#microsoft.graph.aadUserConversationMember", "roles": ["owner"],
                     "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{me['id']}')"},
                    {"@odata.type": "#microsoft.graph.aadUserConversationMember", "roles": ["owner"],
                     "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{tgt['id']}')"},
                ],
            })
            msg = graph(t, "POST", f"/chats/{chat['id']}/messages", {
                "body": {"contentType": "html", "content": html_body},
            })
            print(f"  [{i:2d}/{len(targets)}] OK -> {tgt['displayName']} <{upn}>")
            log.append({"upn": upn, "name": tgt.get("displayName"),
                        "ok": True, "chat_id": chat["id"], "message_id": msg.get("id")})
        except urllib.error.HTTPError as e:
            print(f"  [{i:2d}/{len(targets)}] NG -> {upn}: HTTP {e.code} {getattr(e, 'body', '')[:200]}")
            log.append({"upn": upn, "ok": False, "error": f"HTTP {e.code}"})
        except Exception as e:
            print(f"  [{i:2d}/{len(targets)}] NG -> {upn}: {e}")
            log.append({"upn": upn, "ok": False, "error": str(e)})
        time.sleep(0.8)

    LOG.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    ok = sum(1 for x in log if x.get("ok"))
    print(f"\nDONE: {ok}/{len(targets)} sent -> {LOG.name}")


if __name__ == "__main__":
    main()
