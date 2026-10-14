"""パレットラベルの印刷係。
GAS の印刷待ち（確定したパレット）を数秒ごとに取りに行き、担当のプリンターへ Wi-Fi で送る。
プリンターは MAC アドレスで指定し、IP はその場で探す（DHCP で変わってもよいように）。

  label-printer.exe（python 印刷係.py）  … 黒い画面つきで常駐して印刷する（閉じると止まる。記録は 印刷係.log にも残る）
  label-printer.exe --install           … ログイン時の自動起動を設定して、今すぐ起動する
  label-printer.exe --uninstall         … 自動起動をやめて止める
  label-printer.exe --status            … 動いているかと最近の記録
  label-printer.exe --stop              … 止める
  label-printer.exe --test              … 既定のプリンターに見本を1枚出す
  label-printer.exe --qr [MAC]          … そのプリンターに、本体に貼る登録用QRを出す
画面が出ていないと止まっているように見えて現場が心配するので、常駐も画面つきにしている。
exe の名前を英数字にしているのは、bat に日本語を書くとコマンドプロンプトの文字コード次第で化けるため。
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

# exe にした時は exe の隣、python で動かす時はこのファイルの隣に 設定.json を置く
HERE = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
CONF = os.path.join(HERE, "設定.json")
DONE = os.path.join(HERE, "印刷済み.txt")
LOG = os.path.join(HERE, "印刷係.log")
PORT = 9100


def log(msg):
    line = dt.datetime.now().strftime("%m/%d %H:%M:%S ") + msg
    if sys.stdout:          # 画面なし版では stdout が無い
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


# ===== Wi-Fi を見張る =====
# 設定.json の "wifi" に名前があれば、それ以外につながっていたら切り替える（無ければ何もしない）
def current_ssid():
    out = subprocess.run(["netsh", "wlan", "show", "interfaces"], capture_output=True, text=True, errors="replace").stdout
    m = re.search(r"^\s*SSID\s*:\s*(.+?)\s*$", out, re.M)    # BSSID の行は頭に B が付くので当たらない
    return m.group(1) if m else ""


class WifiGuard:
    EVERY = 30

    def __init__(self, conf):
        self.want = conf.get("wifi", "")
        self.at = 0

    def check(self):
        if not self.want or time.time() - self.at < self.EVERY:
            return
        self.at = time.time()
        now = current_ssid()
        if now == self.want:
            return
        log(f"Wi-Fi が {now or '未接続'} なので {self.want} につなぎ直します")
        r = subprocess.run(["netsh", "wlan", "connect", f"name={self.want}"], capture_output=True, text=True, errors="replace")
        for _ in range(15):
            time.sleep(1)
            if current_ssid() == self.want:
                log(f"{self.want} につながりました")
                return
        log(f"{self.want} につなげませんでした（この PC で一度手でつないでパスワードを保存しておく） {r.stdout.strip()}")


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
        "reprint": bool(job.get("reprint")),
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


MUTEX = r"Local\PalletLabelPrinter"


def set_title(text):
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(text)
    except Exception:
        pass


def only_one():
    """2つ目が起動したら False。同じ依頼を2つの印刷係が同時に出して二重になるのを防ぐ"""
    import ctypes
    k = ctypes.windll.kernel32
    only_one.handle = k.CreateMutexW(None, False, MUTEX)
    return k.GetLastError() != 183    # ERROR_ALREADY_EXISTS


def run():
    if not only_one():
        log("印刷係はすでに動いています（2つ目は起動しません）")
        time.sleep(5)
        return
    conf = load_conf()
    finder = Finder()
    wifi = WifiGuard(conf)
    done = load_done()        # 印刷したのに「印刷済み」を返せなかった分を二度出さないため
    unsent = set()
    errors = {}
    if sys.stdout:
        print("=" * 56)
        print("  パレットラベルの印刷係です。")
        print("  この画面を閉じると、ラベルの印刷が止まります。")
        print("  （最小化は大丈夫です）")
        print("=" * 56)
    log("印刷係を開始しました")
    while True:
        set_title("印刷係 - 動いています（最終確認 " + dt.datetime.now().strftime("%H:%M:%S") + "）  閉じると止まります")
        try:
            wifi.check()
            if unsent:
                api(conf, {"action": "printed", "pallets": sorted(unsent)})
                unsent.clear()
            r = api(conf, {"action": "printQueue"})
            if not r.get("ok"):
                log(f"GAS がエラーを返しました: {r.get('error')}（鍵・URL を確認）")
                time.sleep(60)
                continue
            for job in r.get("jobs", []):
                p = job.get("id") or job["pallet"]     # 再発行は パレット番号-R時刻 で別に来る
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
                    log(f"○ {p}{'（再発行）' if job.get('reprint') else ''}  {job.get('carrier')}  {job.get('count')}箱  {job.get('name') or job.get('user')}")
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


def qr_tag():
    i = sys.argv.index("--qr")
    mac = sys.argv[i + 1] if len(sys.argv) > i + 1 else input("プリンターのMACアドレス（例 00:11:22:AA:BB:CC）: ")
    if not norm_mac(mac):
        log("MACアドレスの形が違います")
        return
    ip = Finder().find(mac)
    if not ip:
        log(f"プリンター {norm_mac(mac)} が見つかりません（電源・Wi-Fi を確認）")
        return
    send(ip, ラベル.printer_tag(mac)[1])
    log(f"{norm_mac(mac)} に登録用QRを出しました")


# ===== 自動起動まわり（bat から呼ぶ） =====
AGENT = "label-printer.exe"   # 常駐も操作も同じ exe。区別は引数の有無
LNK_PATH = "(Join-Path ([Environment]::GetFolderPath('Startup')) '印刷係.lnk')"   # PowerShell の式


def _ps(script):
    """日本語を含む PowerShell をそのまま渡す（-EncodedCommand は UTF-16 なので化けない）"""
    import base64
    enc = base64.b64encode(script.encode("utf-16-le")).decode()
    return subprocess.run(["powershell", "-NoProfile", "-EncodedCommand", enc],
                          capture_output=True, text=True, errors="replace")


def _running():
    """常駐している印刷係がいるか。同じ exe を --status で動かしている自分を数えないよう、Mutex で見る"""
    import ctypes
    h = ctypes.windll.kernel32.OpenMutexW(0x00100000, False, MUTEX)   # SYNCHRONIZE
    if h:
        ctypes.windll.kernel32.CloseHandle(h)
    return bool(h)


def install():
    exe = os.path.join(HERE, AGENT)
    if not os.path.exists(exe):
        print(f"{AGENT} が同じフォルダにありません")
        return
    r = _ps(f"""$s = (New-Object -ComObject WScript.Shell).CreateShortcut({LNK_PATH})
$s.TargetPath = '{exe}'
$s.WorkingDirectory = '{HERE}'
$s.Save()""")
    if r.returncode:
        print("自動起動を設定できませんでした: " + r.stderr.strip())
        return
    if not _running():
        subprocess.Popen([exe], cwd=HERE, creationflags=0x00000010)   # CREATE_NEW_CONSOLE（黒い画面を別に開く）
        time.sleep(3)
    print("自動起動を設定しました。次から Windows にログインすると、印刷係の黒い画面が自動で開きます。")
    want = load_conf().get("wifi", "")
    if want:
        r = subprocess.run(["netsh", "wlan", "set", "profileparameter", f"name={want}", "connectionmode=auto"],
                           capture_output=True, text=True, errors="replace")
        if r.returncode == 0:
            print(f"Wi-Fi「{want}」を自動で接続するようにしました。")
        else:
            print(f"Wi-Fi「{want}」の自動接続は設定できませんでした（{r.stdout.strip() or r.stderr.strip()}）。")
            print(f"　Wi-Fi の一覧で {want} を選び、「自動的に接続」にチェックを入れてください。")
        print(f"ほかの Wi-Fi につながっていたら、印刷係が 30 秒以内に {want} へ切り替えます。")
    status()


def stop_agent():
    if not _running():
        print("印刷係は動いていませんでした")
        return
    # 引数なしで動いている（＝常駐している）ものだけ止める。--stop を実行中の自分は止めない
    _ps(f"Get-CimInstance Win32_Process -Filter \"Name='{AGENT}'\" | Where-Object {{ $_.CommandLine -notmatch ' --' }} | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }")
    print("印刷係を止めました")


def uninstall():
    _ps(f"Remove-Item -ErrorAction SilentlyContinue {LNK_PATH}")
    stop_agent()
    print("自動起動をやめました")


def status():
    print("○ 印刷係は動いています" if _running() else "× 印刷係は止まっています（「1_最初に1回だけ押す」の bat で起動できます）")
    auto = _ps(f"Test-Path {LNK_PATH}").stdout.strip() == "True"
    print("　自動起動: " + ("設定済み" if auto else "なし"))
    print()
    print("----- 最近の記録 -----")
    if os.path.exists(LOG):
        with open(LOG, encoding="utf-8-sig", errors="replace") as f:
            for line in f.readlines()[-15:]:
                print(line.rstrip())


if __name__ == "__main__":
    try:
        if "--test" in sys.argv:
            test()
        elif "--qr" in sys.argv:
            qr_tag()
        elif "--install" in sys.argv:
            install()
        elif "--uninstall" in sys.argv:
            uninstall()
        elif "--status" in sys.argv:
            status()
        elif "--stop" in sys.argv:
            stop_agent()
        else:
            run()
    except Exception as e:      # exe をダブルクリックした時に、理由が読めないまま閉じないように
        log(f"止まりました: {e}")
        if sys.stdin:
            input("Enter で閉じます")
