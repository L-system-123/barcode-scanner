"""パレットラベル（55×86mm, 203dpi）を SBPL にする。
文字は画像にして送り（漢字フォントの有無に左右されない）、QR はプリンターに描かせる。
"""
from PIL import Image, ImageDraw, ImageFont

W = 384          # 印字幅 48mm（8ドット/mm）
H = 640          # 80mm（用紙86mm、上下に余白）
FONT_B = r"C:\Windows\Fonts\BIZ-UDGothicB.ttc"
FONT_R = r"C:\Windows\Fonts\BIZ-UDGothicR.ttc"
QR_X, QR_Y, QR_CELL = 8, 168, 7


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
    return img


def to_sbpl(img, qr_data):
    w8, h8 = (img.width + 7) // 8, (img.height + 7) // 8
    bits = Image.new("1", (w8 * 8, h8 * 8), 1)
    bits.paste(img, (0, 0))
    raw = bytes(b ^ 0xFF for b in bits.tobytes())   # PIL: 1=白 → SBPL: 1=黒
    E = b"\x1b"
    return (E + b"A"
            + E + b"V0001" + E + b"H0001"
            + E + b"GH%03d%03d" % (w8, h8) + raw.hex().upper().encode()
            + E + b"V%04d" % QR_Y + E + b"H%04d" % QR_X
            + E + b"2D30,M,%02d,0,0" % QR_CELL + E + b"DS2," + qr_data.encode("ascii")
            + E + b"Q1" + E + b"Z")


def make(p):
    img = render(p)
    return img, to_sbpl(img, p["pallet_no"])
