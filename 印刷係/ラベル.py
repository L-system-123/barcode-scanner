"""パレットラベル（55×86mm, 203dpi）を SBPL にする。
文字は画像にして送り（漢字フォントの有無に左右されない）、QR はプリンターに描かせる。
"""
import os

from PIL import Image, ImageDraw, ImageFont

W = 384          # 印字幅 48mm（8ドット/mm）
H = 640          # 80mm（用紙86mm、上下に余白）
QR_X, QR_Y, QR_CELL = 8, 168, 7

# BIZ UDゴシックが無い古いWindowsでも出るように、メイリオ・MSゴシックへ順に逃がす
_FONTS = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
FONT_B = next((os.path.join(_FONTS, f) for f in ("BIZ-UDGothicB.ttc", "meiryob.ttc", "msgothic.ttc")
               if os.path.exists(os.path.join(_FONTS, f))), "msgothic.ttc")
FONT_R = next((os.path.join(_FONTS, f) for f in ("BIZ-UDGothicR.ttc", "meiryo.ttc", "msgothic.ttc")
               if os.path.exists(os.path.join(_FONTS, f))), "msgothic.ttc")


def _font(path, size):
    return ImageFont.truetype(path, size)


def _fit(draw, text, path, size, max_w):
    """幅に収まるまで文字を小さくする"""
    while size > 16:
        f = _font(path, size)
        if draw.textlength(text, font=f) <= max_w:
            return f
        size -= 2
    return _font(path, size)


def render(p):
    """p: dict(carrier, dest_cd, boxes, pallet_no, time, user) → 1bit画像"""
    img = Image.new("1", (W, H), 1)
    d = ImageDraw.Draw(img)

    f = _fit(d, p["carrier"], FONT_B, 56, W - 16)
    d.text((8, 8), p["carrier"], font=f, fill=0)
    d.text((8, 84), "配送CD " + p["dest_cd"], font=_font(FONT_B, 40), fill=0)
    d.line((0, 148, W, 148), fill=0, width=3)

    # QR（プリンター側で描く）の右に箱数
    x = 216
    d.text((x, 172), "箱数", font=_font(FONT_R, 30), fill=0)
    n = str(p["boxes"])
    f = _fit(d, n, FONT_B, 120, W - x - 4)
    d.text((x, 212), n, font=f, fill=0)

    d.line((0, 380, W, 380), fill=0, width=3)
    f = _fit(d, p["pallet_no"], FONT_B, 36, W - 16)
    d.text((8, 396), p["pallet_no"], font=f, fill=0)
    d.text((8, 452), p["time"], font=_font(FONT_R, 30), fill=0)
    d.text((8, 496), "担当 " + p["user"], font=_fit(d, "担当 " + p["user"], FONT_R, 30, W - 16), fill=0)
    if p.get("reprint"):
        f = _font(FONT_B, 32)
        tw = d.textlength("再発行", font=f)
        d.rectangle((W - tw - 24, 444, W - 4, 488), fill=0)
        d.text((W - tw - 14, 450), "再発行", font=f, fill=1)
    return img


def to_sbpl(img, qr_data, qr_x=QR_X, qr_y=QR_Y, qr_cell=QR_CELL):
    w8, h8 = (img.width + 7) // 8, (img.height + 7) // 8
    bits = Image.new("1", (w8 * 8, h8 * 8), 1)
    bits.paste(img, (0, 0))
    raw = bytes(b ^ 0xFF for b in bits.tobytes())   # PIL: 1=白 → SBPL: 1=黒
    E = b"\x1b"
    return (E + b"A"
            + E + b"V0001" + E + b"H0001"
            + E + b"GH%03d%03d" % (w8, h8) + raw.hex().upper().encode()
            + E + b"V%04d" % qr_y + E + b"H%04d" % qr_x
            + E + b"2D30,M,%02d,0,0" % qr_cell + E + b"DS2," + qr_data.encode("ascii")
            + E + b"Q1" + E + b"Z")


def make(p):
    img = render(p)
    return img, to_sbpl(img, p["pallet_no"])


def printer_tag(mac):
    """プリンター本体に貼るシール。iPhone の「プリンターを登録」でこのQRを読む"""
    hexmac = "".join(c for c in mac.upper() if c in "0123456789ABCDEF")
    show = ":".join(hexmac[i:i + 2] for i in range(0, 12, 2))
    img = Image.new("1", (W, H), 1)
    d = ImageDraw.Draw(img)
    d.text((8, 8), "プリンター", font=_font(FONT_B, 48), fill=0)
    d.text((8, 470), show, font=_fit(d, show, FONT_B, 40, W - 16), fill=0)
    d.text((8, 530), "スキャン画面の「プリンターを登録」で読む", font=_fit(d, "スキャン画面の「プリンターを登録」で読む", FONT_R, 24, W - 16), fill=0)
    # QR の中身は16文字 → 21セル × 13ドット ≒ 270ドット。上下の文字の間に入る
    return img, to_sbpl(img, "PRN:" + hexmac, qr_x=40, qr_y=80, qr_cell=13)
