"""GitHub worker — bitta ko'chani GitHub IP'dan yuritadi.

Lokal bridge'dan baza/IIB/state oladi, saytda ishlaydi, natijalarni
jonli ravishda (har 20 s) qaytarib turadi. Baza/hisobot lokalda qoladi.
"""
import base64
import json
import os
import subprocess
import sys
import threading
import time
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
BRIDGE = os.environ["BRIDGE_URL"].rstrip("/")
BUSER = os.environ["BRIDGE_USER"]
BPASS = os.environ["BRIDGE_PASS"]
STREET = os.environ["STREET_ID"]
SNAME = os.environ.get("STREET_NAME", STREET)
SLOT = os.environ.get("SLOT", "GH")
TOKEN = os.environ["AUTH_TOKEN"]
HOUSES = os.environ.get("HOUSES", "0")

BASE = "/tmp/base.xlsx"
IIB = "/tmp/iib.xlsx"
REPORT = "/tmp/report.jsonl"
STATE = os.path.join(HERE, f"state_v2_{STREET}.json")
RUNLOG = os.path.join(HERE, "run_new_1.log")

STOP = threading.Event()
OFFSET = 0


def auth_header():
    raw = f"{BUSER}:{BPASS}".encode("utf-8")
    return "Basic " + base64.b64encode(raw).decode("ascii")


def get(path, out=None, timeout=300):
    req = urllib.request.Request(BRIDGE + path, headers={"Authorization": auth_header()})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    if out:
        with open(out, "wb") as f:
            f.write(data)
        return len(data)
    return data


def post(path, body, timeout=180):
    if isinstance(body, str):
        body = body.encode("utf-8")
    req = urllib.request.Request(BRIDGE + path, data=body, method="POST",
                                 headers={"Authorization": auth_header(),
                                          "Content-Type": "application/octet-stream"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()[:200]


def flush_report():
    """Yangi JSONL qatorlarini bridge'ga yuboradi (hisobot darhol yoziladi)."""
    global OFFSET
    if not os.path.exists(REPORT):
        return
    size = os.path.getsize(REPORT)
    if size <= OFFSET:
        return
    with open(REPORT, "rb") as f:
        f.seek(OFFSET)
        data = f.read(size - OFFSET)
    cut = data.rfind(b"\n")
    if cut < 0:
        return
    chunk = data[:cut + 1]
    q = urllib.parse.urlencode({"street": STREET, "name": SNAME})
    try:
        post("/report?" + q, chunk)
        OFFSET += len(chunk)
    except Exception as e:
        print("report yuborilmadi:", str(e)[:120], flush=True)


def push_log():
    if not os.path.exists(RUNLOG):
        return
    try:
        with open(RUNLOG, "rb") as f:
            f.seek(0, 2)
            size = f.tell()
            f.seek(max(0, size - 30000))
            txt = f.read()
        i = txt.find(b"\n")
        if i > 0:
            txt = txt[i + 1:]
        q = urllib.parse.urlencode({"slot": SLOT, "street": STREET, "name": SNAME})
        post("/log?" + q, txt)
    except Exception as e:
        print("log yuborilmadi:", str(e)[:120], flush=True)


RUN_ID = os.environ.get("GITHUB_RUN_ID", "")


def publish_vnc():
    url = os.environ.get("VNC_URL", "")
    if not url:
        return
    try:
        post("/vnc?slot=" + SLOT,
             json.dumps({"url": url, "pass": os.environ.get("VNC_PASS", ""),
                         "run": RUN_ID}))
    except Exception:
        pass


def heartbeat():
    while not STOP.is_set():
        flush_report()
        push_log()
        publish_vnc()
        try:
            if os.path.exists(STATE):
                with open(STATE, "rb") as f:
                    post("/state?street=" + STREET, f.read(), timeout=30)
        except Exception:
            pass
        STOP.wait(20)


def screenshots():
    """Eski screenshot oqimi — endi VNC ishlatiladi (o'chirilgan)."""
    return


def wait_cdp(url, seconds=40):
    for _ in range(seconds * 2):
        try:
            with urllib.request.urlopen(url + "/json/version", timeout=2):
                return True
        except Exception:
            time.sleep(0.5)
    return False


def main():
    print("bridge:", BRIDGE)
    print("street:", STREET, SNAME, "slot:", SLOT)
    vnc_url = os.environ.get("VNC_URL", "")
    if vnc_url:
        publish_vnc()
        print("VNC e'lon qilindi", flush=True)
    print("base yuklanmoqda:", get("/base.xlsx", BASE))
    print("iib yuklanmoqda:", get("/iib.xlsx", IIB, timeout=1200))
    try:
        st = get("/state?street=" + STREET)
        if st and st.strip() not in (b"", b"{}"):
            with open(STATE, "wb") as f:
                f.write(st)
            print("state olindi:", len(st))
    except Exception as e:
        print("state olinmadi:", e)

    if os.path.exists(REPORT):
        os.remove(REPORT)

    from playwright.sync_api import sync_playwright

    global OFFSET
    OFFSET = 0

    with sync_playwright() as p:
        exe = p.chromium.executable_path
        args = [exe, "--remote-debugging-port=9222",
                "--user-data-dir=/tmp/ghprof", "--no-sandbox", "--disable-dev-shm-usage",
                "--no-first-run", "--no-default-browser-check",
                "--window-size=1440,880", "--window-position=0,0"]
        if os.environ.get("HEADLESS") == "1":
            args.append("--headless=new")
        proc = subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            if not wait_cdp("http://127.0.0.1:9222"):
                print("CDP ochilmadi")
                return 2
            b = p.chromium.connect_over_cdp("http://127.0.0.1:9222")
            ctx = b.contexts[0] if b.contexts else b.new_context()
            pg = ctx.new_page()
            pg.goto("https://online-mahalla.uz/", wait_until="domcontentloaded", timeout=90000)
            pg.evaluate(
                """(t) => {
                     localStorage.setItem('auth_token', t);
                     localStorage.setItem('refresh_token', t);
                     localStorage.setItem('auth_user', '{"id":1}');
                   }""", TOKEN)
            r = ctx.request.get("https://api.online-mahalla.uz/auth/user_info",
                                headers={"authorization": "Bearer " + TOKEN}, timeout=30000)
            print("user_info:", r.status)
            # MUHIM: b.close() CDP brauzerni o'chiradi — ulanish ochiq qolishi kerak
        except Exception as e:
            print("browser tayyorlash xatosi:", str(e)[:200])

        env = dict(os.environ)
        env.update({
            "MAHALLA_BASE": BASE,
            "MAHALLA_IIB": IIB,
            "REPORT_JSONL": REPORT,
        })
        cmd = [sys.executable, "mahalla_v2.py", "--street", STREET, "--slot", "1",
               "--cdp", "http://127.0.0.1:9222", "--continue-on-new"]
        if HOUSES and HOUSES != "0":
            cmd += ["--houses", HOUSES]
        print("bot ishga tushdi:", " ".join(cmd[1:]), flush=True)

        th = threading.Thread(target=heartbeat, daemon=True)
        th.start()
        out = open("/tmp/botout.log", "a", encoding="utf-8")
        t0 = time.time()
        rc = 0
        for attempt in range(1, 6):
            out = open("/tmp/botout.log", "a", encoding="utf-8")
            rc = subprocess.run(cmd, cwd=HERE, env=env, stdout=out,
                                stderr=subprocess.STDOUT).returncode
            out.close()
            try:
                txt = open(RUNLOG, encoding="utf-8", errors="replace").read()
            except Exception:
                txt = ""
            if "TUGADI:" in txt or "LOGIN KERAK" in txt:
                break
            print(f"bot to'xtadi (urinish {attempt}/5) — 10 s dan keyin davom", flush=True)
            STOP.wait(10)
        dur = time.time() - t0
        STOP.set()
        th.join(timeout=15)
        print("bot rc:", rc, "davomiylik:", round(dur, 1), "s")

        flush_report()
        push_log()
        if dur < 10 and os.path.exists("/tmp/botout.log"):
            with open("/tmp/botout.log", "rb") as f:
                err = f.read()[-8000:]
            q = urllib.parse.urlencode({"slot": SLOT, "street": STREET, "name": SNAME})
            try:
                post("/log?" + q, err)
            except Exception:
                pass
            print("bot juda tez to'xtadi — log lokal bridge'ga yuborildi")

        if os.path.exists(STATE):
            with open(STATE, "rb") as f:
                post("/state?street=" + STREET, f.read())

        ok = fail = skip = 0
        if os.path.exists(REPORT):
            for line in open(REPORT, encoding="utf-8"):
                try:
                    d = json.loads(line)
                except Exception:
                    continue
                if d.get("ev") == "ok":
                    ok += 1
                else:
                    fail += 1
        print(f"WORKER_DONE ok={ok} fail={fail} skip={skip}", flush=True)
        if vnc_url:
            try:
                post("/vnc?slot=" + SLOT,
                     json.dumps({"url": "off", "pass": "", "run": RUN_ID}))
            except Exception:
                pass
        proc.terminate()
        return 0


if __name__ == "__main__":
    sys.exit(main())
