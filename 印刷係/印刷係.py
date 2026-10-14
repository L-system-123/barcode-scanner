"""パレットラベルの印刷係。
GAS の印刷待ち（確定したパレット）を数秒ごとに取りに行き、担当のプリンターへ Wi-Fi で送る。
プリンターは MAC アドレスで指定し、IP はその場で探す（DHCP で変わってもよいように）。

  python 印刷係.py            … 常駐して印刷する
  python 印刷係.py --test     … 既定のプリンターに見本を1枚出す
設定は同じフォルダの 設定.json（公開しない）。
"""
import concurrent.futures as cf
import datetime as dt
import ipaddress
import json
import os
import re
import socket
import subprocess
import sys
import time
import urllib.request

import ラベル

HERE = os.path.dirname(os.path.abspath(__file__))
CONF = os.path.join(HERE, "設定.json")
DONE = os.path.join(HERE, "印刷済み.txt")
LOG = os.path.join(HERE, "印刷係.log")
PORT = 9100


def log(msg):
    line = dt.datetime.now().strftime("%m/%d %H:%M:%S ") + msg
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8-sig") as f:
        f.write(line + "\n")


def load_conf():
    with open(CONF, encoding="utf-8-sig") as f:
        c = json.load(f)
    c.setdefault("poll_sec", 5)
    return c


# ===== GAS =====
def api(conf, body):
    body = dict(body, key=conf["key"])
    req = urllib.request.Request(conf["api"], data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.loads(r.read().decode("utf-8"))


# ===== プリンターを MAC から探す =====
def norm_mac(s):
    h = re.sub(r"[^0-9a-f]", "", str(s).lower())
    return ":".join(h[i:i + 2] for i in range(0, 12, 2)) if len(h) == 12 else ""


def arp_table():
    out = subprocess.run(["arp", "-a"], capture_output=True, text=True, errors="replace").stdout
    t = {}
    for ip, mac in re.findall(r"(\d+\.\d+\.\d+\.\d+)\s+([0-9a-fA-F]{2}(?:[-:][0-9a-fA-F]{2}){5})", out):
        t[norm_mac(mac)] = ip
    return t


def my_network():
    """今つながっているネットワーク（例 192.168.1.0/24）"""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    finally:
        s.close()
    ps = subprocess.run(["powershell", "-NoProfile", "-Command",
                         f"(Get-NetIPAddress -IPAddress {ip}).PrefixLength"],
                        capture_output=True, text=True).stdout.strip()
    prefix = max(int(ps or 24), 22)   # 広すぎるネットワークは /22（1,000台分）までしか探さない
    return ipaddress.ip_network(f"{ip}/{prefix}", strict=False)


def _touch(ip):
    """つなぎに行くだけ。返事がなくても ARP 表に MAC が載る"""
    try:
        socket.create_connection((str(ip), PORT), timeout=0.4).close()
    except OSError:
        pass


class Finder:
    def __init__(self):
        self.ip = {}

    def find(self, mac, fresh=False):
        mac = norm_mac(mac)
        if not mac:
            return None
        if not fresh and mac in self.ip:
            return self.ip[mac]
        hit = arp_table().get(mac)
        if not hit:
            net = my_network()
            log(f"プリンター {mac} を {net} の中で探しています…")
            with cf.ThreadPoolExecutor(128) as ex:
                list(ex.map(_touch, net.hosts()))
            hit = arp_table().get(mac)
        if hit:
            if self.ip.get(mac) != hit:
                log(f"プリンター {mac} → {hit}")
            self.ip[mac] = hit
        return hit


def send(ip, data):
    with socket.create_connection((ip, PORT), timeout=10) as s:
        s.sendall(data)


# ===== 1枚出す =====
def label_of(job):
    at = dt.datetime.fromtimestamp(job["at"] / 1000) if job.get("at") else dt.datetime.now()
    return {
        "carrier": job.get("carrier", ""),
        "dest_cd": job.get("cd", "") or "－",
        "boxes": job.get("count", 0),
        "pallet_no": job["pallet"],
        "time": at.strftime("%Y/%m/%d %H:%M"),
        "user": job.get("name") or job.get("user", ""),
    }


def print_job(finder, conf, job):
    mac = job.get("printer") or conf.get("default_printer", "")
    if not norm_mac(mac):
        raise RuntimeError(f"プリンターが決まっていません（担当 {job.get('user')}。ユーザー タブの H列か 設定.json の default_printer）")
    _, data = ラベル.make(label_of(job))
    ip = finder.find(mac)
    try:
        if not ip:
            raise OSError("見つからない")
        send(ip, data)
    except OSError:
        ip = finder.find(mac, fresh=True)   # IP が変わったかもしれないので探し直して1回だけやり直す
        if not ip:
            raise RuntimeError(f"プリンター {norm_mac(mac)} が見つかりません（電源・Wi-Fi を確認）")
        send(ip, data)


def load_done():
    today = dt.date.today().strftime("%Y%m%d")
    if not os.path.exists(DONE):
        return set()
    with open(DONE, encoding="utf-8-sig") as f:
        return {x.strip() for x in f if x.startswith(today)}


def run():
    conf = load_conf()
    finder = Finder()
    done = load_done()        # 印刷したのに「印刷済み」を返せなかった分を二度出さないため
    unsent = set()
    errors = {}
    log("印刷係を開始しました（止めるときはこの画面を閉じる）")
    while True:
        try:
            if unsent:
                api(conf, {"action": "printed", "pallets": sorted(unsent)})
                unsent.clear()
            r = api(conf, {"action": "printQueue"})
            if not r.get("ok"):
                log(f"GAS がエラーを返しました: {r.get('error')}（鍵・URL を確認）")
                time.sleep(60)
                continue
            for job in r.get("jobs", []):
                p = job["pallet"]
                if p not in done:
                    try:
                        print_job(finder, conf, job)
                    except Exception as e:
                        if errors.get(p) != str(e):      # 同じエラーを毎回書かない
                            log(f"× {p} 印刷できません: {e}")
                        errors[p] = str(e)
                        continue
                    done.add(p)
                    errors.pop(p, None)
                    with open(DONE, "a", encoding="utf-8-sig") as f:
                        f.write(p + "\n")
                    log(f"○ {p}  {job.get('carrier')}  {job.get('count')}箱  {job.get('name') or job.get('user')}")
                unsent.add(p)
            if unsent:
                api(conf, {"action": "printed", "pallets": sorted(unsent)})
                unsent.clear()
        except Exception as e:
            log(f"通信できません: {e}")
            time.sleep(15)
        time.sleep(conf["poll_sec"])


def test():
    conf = load_conf()
    job = {"pallet": dt.date.today().strftime("%Y%m%d") + "-TEST-001", "carrier": "サンプル運送", "cd": "0841",
           "count": 24, "at": time.time() * 1000, "user": "TEST", "name": "テスト担当", "printer": ""}
    print_job(Finder(), conf, job)
    log("見本を1枚送りました")


if __name__ == "__main__":
    test() if "--test" in sys.argv else run()
