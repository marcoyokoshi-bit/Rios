"""Gerador de fluxograma PNG (vertical, caixas arredondadas + setas) a partir das etapas de um POP."""
import os
from PIL import Image, ImageDraw, ImageFont


def _find_font(bold=True):
    candidates_win = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
    ]
    candidates_linux = [
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    candidates_mac = [
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for path in candidates_win + candidates_linux + candidates_mac:
        if os.path.exists(path):
            return path
    return None


def _load_font(size, bold=True):
    path = _find_font(bold)
    if path:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def hex_to_rgb(hex_color):
    h = (hex_color or "5E4778").lstrip('#')
    if len(h) != 6:
        h = "5E4778"
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def _mix(c1, c2, t):
    return tuple(int(c1[i] + (c2[i]-c1[i])*t) for i in range(3))


def build_shades(base_hex, n=5):
    """Gera até n variações coerentes da cor institucional (mais escura -> mais clara)."""
    base = hex_to_rgb(base_hex)
    white = (255, 255, 255)
    black = (20, 20, 24)
    stops = []
    # 2 tons mais escuros, o tom base, 2 tons mais claros (mas ainda com contraste p/ texto branco)
    factors = [-0.28, -0.12, 0.0, 0.16, 0.30]
    for f in factors[:n]:
        if f < 0:
            stops.append(_mix(base, black, -f))
        else:
            stops.append(_mix(base, white, f))
    return stops


def _wrap_text(draw, text, font, max_width):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2]-bbox[0] <= max_width or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def generate_flow_png(etapas, color_hex="5E4778", out_path="flow.png", width=1600):
    """etapas: lista de dicts com 'titulo'. Retorna (out_path, width_px, height_px)."""
    titulos = [ (e.get('titulo') or f"Etapa {i+1}").strip().upper() for i, e in enumerate(etapas) ]
    n = max(len(titulos), 1)
    palette = build_shades(color_hex, 5)

    box_w = int(width * 0.60)
    box_h = 140
    gap = 74
    pad = 50
    height = pad*2 + n*box_h + (n-1)*gap

    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    font = _load_font(34, bold=True)

    x0 = (width - box_w)//2
    y = pad
    arrow_color = (130, 130, 130)
    for i, titulo in enumerate(titulos):
        color = palette[i % len(palette)]
        draw.rounded_rectangle([x0, y, x0+box_w, y+box_h], radius=26, fill=color)
        lines = _wrap_text(draw, titulo, font, box_w - 60)
        line_h = font.size + 8
        total_h = line_h * len(lines)
        ty = y + (box_h - total_h)//2
        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            tw = bbox[2]-bbox[0]
            draw.text((x0 + (box_w-tw)//2, ty), line, font=font, fill="white")
            ty += line_h
        if i < n-1:
            ax = x0 + box_w//2
            ay1 = y + box_h + 6
            ay2 = y + box_h + gap - 22
            draw.line([(ax, ay1), (ax, ay2)], fill=arrow_color, width=6)
            draw.polygon([(ax-16, ay2), (ax+16, ay2), (ax, ay2+22)], fill=arrow_color)
        y += box_h + gap

    img.save(out_path, dpi=(200, 200))
    return out_path, width, height


if __name__ == "__main__":
    etapas = [
        {"titulo": "Aviso e preparação"},
        {"titulo": "Inspeção inicial de qualidade"},
        {"titulo": "Pesagem e conferência"},
        {"titulo": "Aprovação e liberação"},
        {"titulo": "Armazenamento"},
    ]
    p, w, h = generate_flow_png(etapas, "5E4778", "/tmp/rios_docx_dev/flow_test.png")
    print(p, w, h)
