"""Tunnel URL + parolni repo fayliga yozadi (log oqimini kutmaslik uchun)."""
import base64
import json
import os
import urllib.request

repo = os.environ["GITHUB_REPOSITORY"]
tok = os.environ["GH_TOKEN"]
url = os.environ.get("LIVE_URL", "off")
pwd = os.environ.get("LIVE_PASS", "")
text = f"{url}\n{pwd}\n"

api = f"https://api.github.com/repos/{repo}/contents/live_url.txt"
headers = {
    "Authorization": f"Bearer {tok}",
    "Accept": "application/vnd.github+json",
    "X-GitHub-Api-Version": "2022-11-28",
}


def call(method, data=None):
    body = json.dumps(data).encode() if data is not None else None
    req = urllib.request.Request(api, data=body, method=method, headers=headers)
    return urllib.request.urlopen(req, timeout=30).read()


sha = None
try:
    sha = json.loads(call("GET")).get("sha")
except Exception:
    pass

payload = {"message": "live url", "content": base64.b64encode(text.encode()).decode()}
if sha:
    payload["sha"] = sha
try:
    call("PUT", payload)
    print("publish: OK")
except Exception as e:
    print("publish error:", str(e)[:200])
