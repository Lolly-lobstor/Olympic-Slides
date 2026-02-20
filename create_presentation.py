"""
Olympic Hockey Google Slides Presentation Generator
Creates a .pptx file that can be imported into Google Slides
"""

import io
import os
import math
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# ── Slide dimensions (widescreen 16:9) ──────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── Brand colours ────────────────────────────────────────────────────────────
ICE_BLUE    = RGBColor(0x0A, 0x2A, 0x6E)   # deep Olympic blue
GOLD        = RGBColor(0xD4, 0xAF, 0x37)   # gold accent
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE  = RGBColor(0xC8, 0xDF, 0xF5)
DARK_GREY   = RGBColor(0x22, 0x22, 0x22)
RED         = RGBColor(0xC0, 0x20, 0x20)

# ── PIL colour tuples ────────────────────────────────────────────────────────
P_ICE_BLUE  = (10, 42, 110)
P_GOLD      = (212, 175, 55)
P_WHITE     = (255, 255, 255)
P_LIGHT_BLU = (200, 223, 245)
P_DARK_GREY = (34, 34, 34)
P_RED       = (192, 32, 32)
P_ICE_WHITE = (230, 245, 255)


def _pil_buf(img: Image.Image) -> io.BytesIO:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def _draw_text_wrapped(draw, text, x, y, max_w, font, fill, line_h=None):
    """Simple word-wrap text draw."""
    if line_h is None:
        line_h = font.size + 4
    words = text.split()
    line = ""
    cy = y
    for word in words:
        test = (line + " " + word).strip()
        bb = draw.textbbox((0, 0), test, font=font)
        if bb[2] - bb[0] <= max_w:
            line = test
        else:
            if line:
                draw.text((x, cy), line, font=font, fill=fill)
                cy += line_h
            line = word
    if line:
        draw.text((x, cy), line, font=font, fill=fill)


def make_cover_image(w=640, h=480) -> io.BytesIO:
    """Hockey player silhouette + Olympic rings-style image."""
    img = Image.new("RGB", (w, h), P_ICE_BLUE)
    d = ImageDraw.Draw(img)
    # Ice surface gradient effect
    for i in range(h // 2, h):
        alpha = int(200 + 55 * (i - h // 2) / (h // 2))
        d.line([(0, i), (w, i)], fill=(200, 225, 245))
    # Stick shape
    d.rectangle([w//2-8, 80, w//2+8, 340], fill=P_DARK_GREY)
    d.rectangle([w//2-8, 300, w//2+70, 340], fill=P_DARK_GREY)
    # Puck
    d.ellipse([w//2+50, 340, w//2+110, 370], fill=P_DARK_GREY)
    # Olympic rings (5 coloured circles)
    cx, cy, r, gap = w//2-80, 160, 28, 65
    ring_colours = [(0,120,200), (0,0,0), (220,40,40), (255,200,0), (0,160,60)]
    for i, col in enumerate(ring_colours):
        ox = cx + i * gap
        oy = cy if i % 2 == 0 else cy + 20
        d.ellipse([ox-r, oy-r, ox+r, oy+r], outline=col, width=5)
    # Title text
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 36)
        font_sm  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
    except Exception:
        font_big = ImageFont.load_default()
        font_sm  = font_big
    d.text((20, 20), "ICE HOCKEY", font=font_big, fill=P_GOLD)
    d.text((20, 65), "MILAN–CORTINA 2026", font=font_sm, fill=P_WHITE)
    return _pil_buf(img)


def make_history_image(w=480, h=320) -> io.BytesIO:
    """Vintage-style sepia-toned rink illustration."""
    img = Image.new("RGB", (w, h), (180, 140, 90))
    d = ImageDraw.Draw(img)
    # Rink oval
    d.ellipse([30, 30, w-30, h-30], outline=(100, 70, 30), width=5)
    # Center line
    d.line([(w//2, 30), (w//2, h-30)], fill=(100, 70, 30), width=3)
    # Face-off circles
    d.ellipse([w//4-40, h//2-40, w//4+40, h//2+40], outline=(100,70,30), width=2)
    d.ellipse([3*w//4-40, h//2-40, 3*w//4+40, h//2+40], outline=(100,70,30), width=2)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
        font2 = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
    except Exception:
        font = ImageFont.load_default()
        font2 = font
    d.text((w//2-100, 5), "1920 — First Olympic", font=font, fill=(80,50,20))
    d.text((w//2-80, 30), "Hockey Game", font=font2, fill=(80,50,20))
    d.text((10, h-25), "Antwerp, Belgium", font=font2, fill=(80,50,20))
    return _pil_buf(img)


def make_equipment_image(w=480, h=400) -> io.BytesIO:
    """Equipment diagram."""
    img = Image.new("RGB", (w, h), P_ICE_WHITE)
    d = ImageDraw.Draw(img)
    # Helmet
    d.ellipse([w//2-50, 20, w//2+50, 100], fill=(50,50,50))
    d.arc([w//2-50, 60, w//2+50, 120], 0, 180, fill=(80,80,80), width=3)
    # Stick
    d.rectangle([w//2+60, 80, w//2+75, 320], fill=(139,90,43))
    d.rectangle([w//2+60, 290, w//2+130, 320], fill=(139,90,43))
    # Puck
    d.ellipse([w//2+90, 320, w//2+150, 360], fill=P_DARK_GREY)
    # Skate blade
    d.rectangle([w//2-80, 350, w//2+10, 370], fill=(200,200,220))
    d.ellipse([w//2-90, 345, w//2-60, 375], fill=(150,150,170))
    d.ellipse([w//2, 345, w//2+30, 375], fill=(150,150,170))
    # Glove
    d.rounded_rectangle([w//2-140, 180, w//2-60, 260], radius=15, fill=(200,50,50))
    d.rectangle([w//2-130, 150, w//2-70, 200], fill=(180,40,40))
    # Labels
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 16)
    except Exception:
        font = ImageFont.load_default()
    labels = [("HELMET", (w//2-40, 5)), ("STICK", (w//2+135, 180)),
              ("PUCK", (w//2+90, 365)), ("SKATE", (w//2-90, 375)),
              ("GLOVE", (w//2-145, 265))]
    for txt, pos in labels:
        d.text(pos, txt, font=font, fill=P_ICE_BLUE)
    d.text((5, 5), "Hockey Equipment", font=font, fill=P_ICE_BLUE)
    return _pil_buf(img)


def make_skills_image(w=400, h=480) -> io.BytesIO:
    """Athlete silhouette in skating pose."""
    img = Image.new("RGB", (w, h), P_ICE_BLUE)
    d = ImageDraw.Draw(img)
    # Ice
    d.rectangle([0, h*3//4, w, h], fill=P_ICE_WHITE)
    # Body silhouette (skating pose)
    cx = w // 2
    # Head
    d.ellipse([cx-22, 40, cx+22, 84], fill=P_WHITE)
    # Torso (leaning forward)
    d.polygon([(cx-22, 84), (cx+22, 84), (cx+40, 180), (cx-40, 180)], fill=P_LIGHT_BLU)
    # Left arm with stick
    d.line([(cx-22, 100), (cx-60, 170)], fill=P_WHITE, width=10)
    d.line([(cx-60, 170), (cx-20, 280)], fill=(139,90,43), width=6)
    d.line([(cx-20, 280), (cx+20, 280)], fill=(139,90,43), width=6)
    # Right arm
    d.line([(cx+22, 100), (cx+60, 160)], fill=P_WHITE, width=10)
    # Legs
    d.line([(cx-20, 180), (cx-30, 310)], fill=P_LIGHT_BLU, width=14)
    d.line([(cx+20, 180), (cx+50, 280)], fill=P_LIGHT_BLU, width=14)
    # Skates
    d.rectangle([cx-50, 308, cx-10, 322], fill=P_DARK_GREY)
    d.rectangle([cx+30, 278, cx+70, 292], fill=P_DARK_GREY)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
        font_sm = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 14)
    except Exception:
        font = ImageFont.load_default()
        font_sm = font
    d.text((10, 5), "Athlete Skills", font=font, fill=P_GOLD)
    for i, skill in enumerate(["Speed", "Agility", "Strength", "Endurance"]):
        d.text((10, h-90+i*18), f"• {skill}", font=font_sm, fill=P_GOLD)
    return _pil_buf(img)


def make_competition_image(w=640, h=400) -> io.BytesIO:
    """Top-down ice rink diagram."""
    img = Image.new("RGB", (w, h), P_ICE_WHITE)
    d = ImageDraw.Draw(img)
    pad = 20
    # Rink border (rounded rectangle)
    d.rounded_rectangle([pad, pad, w-pad, h-pad], radius=50,
                        outline=P_ICE_BLUE, width=5, fill=(220, 240, 255))
    # Center line
    d.line([(w//2, pad), (w//2, h-pad)], fill=P_RED, width=4)
    # Blue lines
    d.line([(w//3, pad), (w//3, h-pad)], fill=(0,0,180), width=3)
    d.line([(2*w//3, pad), (2*w//3, h-pad)], fill=(0,0,180), width=3)
    # Goalie creases
    d.rectangle([pad, h//2-40, pad+55, h//2+40], fill=(180,210,255), outline=P_RED, width=2)
    d.rectangle([w-pad-55, h//2-40, w-pad, h//2+40], fill=(180,210,255), outline=P_RED, width=2)
    # Center circle
    d.ellipse([w//2-55, h//2-55, w//2+55, h//2+55], outline=P_ICE_BLUE, width=3)
    d.ellipse([w//2-5, h//2-5, w//2+5, h//2+5], fill=P_ICE_BLUE)
    # Face-off circles
    for ox, oy in [(w//4, h//3), (w//4, 2*h//3), (3*w//4, h//3), (3*w//4, 2*h//3)]:
        d.ellipse([ox-30, oy-30, ox+30, oy+30], outline=P_RED, width=2)
        d.ellipse([ox-3, oy-3, ox+3, oy+3], fill=P_RED)
    try:
        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 18)
        font_sm = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    except Exception:
        font = ImageFont.load_default()
        font_sm = font
    d.text((w//2-80, 2), "Olympic Ice Rink", font=font, fill=P_ICE_BLUE)
    d.text((pad+5, h//2-8), "GOAL", font=font_sm, fill=P_RED)
    d.text((w-pad-50, h//2-8), "GOAL", font=font_sm, fill=P_RED)
    d.text((w//2-25, h-18), "200 ft", font=font_sm, fill=P_DARK_GREY)
    return _pil_buf(img)


def make_video_thumb(w=500, h=320) -> io.BytesIO:
    """YouTube-style play button thumbnail."""
    img = Image.new("RGB", (w, h), P_DARK_GREY)
    d = ImageDraw.Draw(img)
    # Background gradient simulation
    for i in range(h):
        c = int(30 + 40 * i / h)
        d.line([(0, i), (w, i)], fill=(c, c+5, c+20))
    # Play button circle
    cx, cy, r = w//2, h//2, 60
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=(180, 0, 0))
    # Triangle
    d.polygon([(cx-20, cy-30), (cx-20, cy+30), (cx+35, cy)], fill=P_WHITE)
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
        font_sm  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
    except Exception:
        font_big = ImageFont.load_default()
        font_sm  = font_big
    d.text((20, 10), "OLYMPIC ICE HOCKEY", font=font_big, fill=P_GOLD)
    d.text((20, 44), "Best Goals & Highlights — Beijing 2022", font=font_sm, fill=P_WHITE)
    d.text((20, h-25), "Duration: ~2 min 45 sec", font=font_sm, fill=P_LIGHT_BLU)
    return _pil_buf(img)


def make_olympian_image(w=400, h=500) -> io.BytesIO:
    """Connor McDavid stylized graphic."""
    img = Image.new("RGB", (w, h), P_ICE_BLUE)
    d = ImageDraw.Draw(img)
    # Jersey number 97
    d.rounded_rectangle([30, 60, w-30, h-60], radius=20, fill=(220, 40, 40))
    d.rounded_rectangle([50, 80, w-50, h-80], radius=15, outline=P_WHITE, width=4)
    try:
        font_huge = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 110)
        font_big  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28)
        font_med  = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 20)
    except Exception:
        font_huge = ImageFont.load_default()
        font_big  = font_huge
        font_med  = font_huge
    # Jersey number
    d.text((w//2-65, 120), "97", font=font_huge, fill=P_WHITE)
    # Name bar
    d.rectangle([30, h-110, w-30, h-60], fill=P_GOLD)
    d.text((50, h-105), "McDavid", font=font_big, fill=P_DARK_GREY)
    d.text((50, h-75), "Connor  |  #97  |  C", font=font_med, fill=P_ICE_BLUE)
    # Top label
    d.text((10, 10), "TEAM CANADA  🍁  2026", font=font_med, fill=P_GOLD)
    return _pil_buf(img)


def make_quiz_image(w=480, h=300) -> io.BytesIO:
    """Olympic rings + quiz banner."""
    img = Image.new("RGB", (w, h), P_ICE_BLUE)
    d = ImageDraw.Draw(img)
    # Olympic rings
    rings = [
        ((60,  80), (0,120,200)),
        ((130, 80), (0,0,0)),
        ((200, 80), (220,40,40)),
        ((270, 80), (255,200,0)),
        ((340, 80), (0,160,60)),
    ]
    for (rx, ry), col in rings:
        shift = 20 if rings.index(((rx,ry),col)) % 2 else 0
        d.ellipse([rx-28, ry-28+shift, rx+28, ry+28+shift], outline=col, width=7)
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 32)
        font_med = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 18)
    except Exception:
        font_big = ImageFont.load_default()
        font_med = font_big
    d.text((30, 160), "FINAL EXAM QUIZ", font=font_big, fill=P_GOLD)
    d.text((30, 205), "Test your Olympic Hockey knowledge!", font=font_med, fill=P_WHITE)
    d.text((30, 235), "Milan–Cortina 2026", font=font_med, fill=P_LIGHT_BLU)
    return _pil_buf(img)


def make_refs_image(w=480, h=300) -> io.BytesIO:
    """References / bibliography graphic."""
    img = Image.new("RGB", (w, h), P_ICE_BLUE)
    d = ImageDraw.Draw(img)
    # Book icon
    d.rounded_rectangle([60, 40, 220, 230], radius=8, fill=P_GOLD)
    d.rounded_rectangle([80, 50, 200, 220], radius=5, fill=P_WHITE)
    for y in range(70, 210, 18):
        d.line([(90, y), (190, y)], fill=P_LIGHT_BLU, width=2)
    d.line([(60, 40), (60, 230)], fill=P_DARK_GREY, width=6)
    # Rings (Olympic rings simplified)
    for i, col in enumerate([(0,120,200),(0,0,0),(220,40,40),(255,200,0),(0,160,60)]):
        ox = 280 + i*35
        d.ellipse([ox-14, 60-14, ox+14, 60+14], outline=col, width=4)
    try:
        font_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28)
        font_med = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 16)
    except Exception:
        font_big = ImageFont.load_default()
        font_med = font_big
    d.text((240, 100), "REFERENCES", font=font_big, fill=P_GOLD)
    d.text((240, 140), "Sources & Citations", font=font_med, fill=P_WHITE)
    d.text((240, 165), "for Olympic Ice Hockey", font=font_med, fill=P_LIGHT_BLU)
    d.text((240, 200), "Presentation", font=font_med, fill=P_LIGHT_BLU)
    return _pil_buf(img)


# Map of image key → generator function
IMAGE_MAKERS = {
    "cover":      make_cover_image,
    "history":    make_history_image,
    "equipment":  make_equipment_image,
    "skills":     make_skills_image,
    "competition":make_competition_image,
    "video":      make_video_thumb,
    "olympian":   make_olympian_image,
    "quiz":       make_quiz_image,
    "refs":       make_refs_image,
}


def set_slide_bg(slide, color: RGBColor):
    """Fill slide background with a solid colour."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    """Add a filled rectangle shape."""
    shape = slide.shapes.add_shape(
        pptx.enum.shapes.MSO_SHAPE_TYPE.AUTO_SHAPE if False else 1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=18, bold=False, color=WHITE,
                 align=PP_ALIGN.LEFT, wrap=True):
    """Add a text box with consistent formatting."""
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txb


def add_title_bar(slide, title_text, subtitle_text=None):
    """Top banner with title and optional subtitle."""
    # Banner rectangle
    add_rect(slide, 0, 0, SLIDE_W, Inches(1.4), ICE_BLUE)
    # Gold accent line
    add_rect(slide, 0, Inches(1.4), SLIDE_W, Inches(0.06), GOLD)
    # Title text
    add_text_box(slide, title_text,
                 Inches(0.3), Inches(0.1), Inches(12.0), Inches(0.8),
                 font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if subtitle_text:
        add_text_box(slide, subtitle_text,
                     Inches(0.3), Inches(0.85), Inches(10.0), Inches(0.45),
                     font_size=16, bold=False, color=GOLD, align=PP_ALIGN.LEFT)


def add_bullet_box(slide, bullets: list[str], left, top, width, height,
                   title=None, title_color=GOLD, bullet_color=DARK_GREY,
                   bg_color=None, font_size=15):
    """Render a list of bullet points inside an optional box."""
    if bg_color:
        add_rect(slide, left, top, width, height, bg_color)
    if title:
        add_text_box(slide, title, left + Inches(0.15), top + Inches(0.1),
                     width - Inches(0.3), Inches(0.45),
                     font_size=17, bold=True, color=title_color, wrap=True)
        y_offset = Inches(0.55)
    else:
        y_offset = Inches(0.1)

    txb = slide.shapes.add_textbox(
        left + Inches(0.15), top + y_offset,
        width - Inches(0.3), height - y_offset - Inches(0.1)
    )
    tf = txb.text_frame
    tf.word_wrap = True
    first = True
    for bullet in bullets:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(4)
        run = p.add_run()
        run.text = f"• {bullet}"
        run.font.size = Pt(font_size)
        run.font.color.rgb = bullet_color


def add_image_safe(slide, img_key: str, left, top, width, height):
    """Generate image with Pillow and embed it."""
    maker = IMAGE_MAKERS.get(img_key)
    if maker:
        try:
            buf = maker()
            slide.shapes.add_picture(buf, left, top, width, height)
            return
        except Exception as e:
            print(f"  [warn] Could not embed image '{img_key}': {e}")
    # Fallback placeholder box
    add_rect(slide, left, top, width, height, LIGHT_BLUE, ICE_BLUE)
    add_text_box(slide, f"[ {img_key.upper()} ]",
                 left, top + height // 2 - Inches(0.2), width, Inches(0.4),
                 font_size=12, bold=True, color=ICE_BLUE, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════════
# SLIDE BUILDERS
# ════════════════════════════════════════════════════════════════════════════

def slide_01_cover(prs):
    """Slide 1 – Cover."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_slide_bg(slide, ICE_BLUE)

    # Diagonal gold banner
    add_rect(slide, 0, Inches(4.8), SLIDE_W, Inches(0.08), GOLD)

    # Main title
    add_text_box(slide, "OLYMPIC ICE HOCKEY",
                 Inches(0.5), Inches(1.0), Inches(9.0), Inches(1.6),
                 font_size=52, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    # Gold underline
    add_rect(slide, Inches(0.5), Inches(2.7), Inches(6.0), Inches(0.07), GOLD)

    add_text_box(slide, "Speed · Skill · Passion on Ice",
                 Inches(0.5), Inches(2.85), Inches(8.0), Inches(0.6),
                 font_size=22, bold=False, color=LIGHT_BLUE, align=PP_ALIGN.LEFT)
    add_text_box(slide, "Milan–Cortina 2026 Winter Olympics",
                 Inches(0.5), Inches(3.5), Inches(8.0), Inches(0.5),
                 font_size=18, bold=False, color=GOLD, align=PP_ALIGN.LEFT)
    add_text_box(slide, "Presented by: Olympic Sports Education Program",
                 Inches(0.5), Inches(5.2), Inches(9.0), Inches(0.45),
                 font_size=14, bold=False, color=LIGHT_BLUE, align=PP_ALIGN.LEFT)
    add_text_box(slide, "February 2026",
                 Inches(0.5), Inches(5.7), Inches(4.0), Inches(0.4),
                 font_size=13, bold=False, color=GOLD, align=PP_ALIGN.LEFT)

    add_image_safe(slide, "cover", Inches(10.0), Inches(1.5), Inches(2.8), Inches(3.5))
    add_image_safe(slide, "quiz",  Inches(9.5),  Inches(5.1), Inches(3.5), Inches(1.8))


def slide_02_history(prs):
    """Slide 2 – History & Background."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "History & Background of Olympic Ice Hockey",
                  "From frozen ponds to the world's biggest stage")

    left_bullets = [
        "Ice hockey originated in Canada in the mid-1800s, with the first recorded indoor game in Montreal in 1875.",
        "Canada dominated early Olympic hockey, winning gold in 1920, 1924, 1928, 1932, and 1948.",
        "Men's ice hockey became an Olympic sport at the 1920 Antwerp Summer Games — the only Summer Olympics to include it.",
        "It became a permanent fixture at the Winter Olympics starting with the 1924 Chamonix Games.",
        "The Soviet Union rose to dominance in the 1950s–1970s, winning multiple Olympic gold medals.",
        "Women's ice hockey was added to the Olympic program at the 1998 Nagano Winter Games.",
        "NHL players were allowed to compete starting with the 1998 Nagano Olympics, raising the level of play dramatically.",
        "The 2026 Milan–Cortina Games continue this storied tradition with the world's best players competing.",
    ]
    add_bullet_box(slide, left_bullets,
                   Inches(0.3), Inches(1.55), Inches(7.8), Inches(5.7),
                   bg_color=LIGHT_BLUE, font_size=13, bullet_color=DARK_GREY)

    add_image_safe(slide, "history", Inches(8.3), Inches(1.6), Inches(4.7), Inches(3.0))

    fun_facts = [
        "First Olympic gold: Canada (1920)",
        "USSR won 7 Olympic golds between 1956–1992",
        "USA's 'Miracle on Ice': 1980 Lake Placid upset over USSR",
        "Canada won gold in both men's and women's in 2002, 2010, and 2014",
    ]
    add_bullet_box(slide, fun_facts,
                   Inches(8.3), Inches(4.8), Inches(4.7), Inches(2.4),
                   title="Key Moments", bg_color=ICE_BLUE,
                   title_color=GOLD, bullet_color=WHITE, font_size=12)


def slide_03_rules(prs):
    """Slide 3 – Rules, Equipment & Scoring."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Rules, Equipment & How Scoring Works",
                  "Everything you need to know to follow the game")

    rules = [
        "Each team has 6 players on ice: 1 goalie, 2 defensemen, 3 forwards.",
        "3 periods × 20 minutes each (regulation time); overtime & shootout if tied.",
        "A goal is scored when the puck fully crosses the goal line into the net.",
        "Offside: a player cannot enter the offensive zone before the puck.",
        "Icing: shooting the puck from behind the center line past the opponent's goal line (without scoring) is illegal.",
        "Penalties: minor (2 min), major (5 min), or misconduct (10 min); team plays shorthanded.",
        "At the Olympics, ties after 3 periods go to a 5-min sudden-death OT, then shootout.",
    ]
    equipment = [
        "Skates – stiff boot with a sharp blade",
        "Stick & puck – vulcanized rubber puck (6 oz)",
        "Helmet with full cage or visor",
        "Shoulder & elbow pads, shin guards",
        "Gloves, padded pants (breezers)",
        "Goalie: extra leg pads, blocker, trapper, mask",
        "Jersey, neck guard, and mouth guard",
    ]

    add_bullet_box(slide, rules, Inches(0.3), Inches(1.55), Inches(6.2), Inches(5.7),
                   title="The Rules", bg_color=ICE_BLUE,
                   title_color=GOLD, bullet_color=WHITE, font_size=12)

    add_bullet_box(slide, equipment, Inches(6.7), Inches(1.55), Inches(3.3), Inches(3.5),
                   title="Equipment", bg_color=LIGHT_BLUE,
                   title_color=ICE_BLUE, bullet_color=DARK_GREY, font_size=12)

    add_image_safe(slide, "equipment", Inches(6.7), Inches(5.15), Inches(3.3), Inches(2.1))
    add_rect(slide, Inches(10.2), Inches(1.55), Inches(2.8), Inches(5.7), LIGHT_BLUE)
    scoring_text = "SCORING\n\nGoal (during play): 1 point\n\nPower Play Goal: 1 point\n(scored while opponent has a penalty)\n\nShort-Handed Goal: 1 point\n(scored while YOUR team has a penalty)\n\nShootout goal: counts only if game is tied after OT"
    add_text_box(slide, scoring_text,
                 Inches(10.3), Inches(1.65), Inches(2.6), Inches(5.5),
                 font_size=11, bold=False, color=DARK_GREY, wrap=True)


def slide_04_skills(prs):
    """Slide 4 – Skills, Fitness, Training & Nutrition."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Skills an Athlete Needs to Compete",
                  "Fitness · Training · Nutrition")

    on_ice = [
        "Skating speed & agility – explosive edge-work and quick turns",
        "Stickhandling – puck control at high speed with both hands",
        "Passing & receiving – crisp, accurate passes under pressure",
        "Shooting – wrist shot, slap shot, backhand with precision",
        "Defensive positioning – gap control, angling opponents",
        "Hockey IQ – reading plays, anticipating puck movement",
    ]
    fitness = [
        "VO₂ max / cardiovascular endurance for 45-sec shifts",
        "Explosive leg power: squats, Olympic lifts, box jumps",
        "Core stability for balance on the ice",
        "Upper body strength for battling in corners & puck battles",
        "Flexibility & agility: lateral movements, hip mobility",
        "Elite players skate up to 5–6 km per game",
    ]
    nutrition = [
        "High carb intake pre-game (pasta, rice, oatmeal) for fuel",
        "Lean protein (chicken, fish, eggs) for muscle repair",
        "Hydration: 2–3 L water/day; electrolytes during play",
        "Recovery: protein shake + carbs within 30 min post-game",
        "Avoid heavy or fried food on game day",
        "Supplements: vitamin D, iron, omega-3 (doctor supervised)",
    ]

    add_bullet_box(slide, on_ice,   Inches(0.2), Inches(1.55), Inches(4.2), Inches(5.7),
                   title="On-Ice Skills", bg_color=ICE_BLUE,
                   title_color=GOLD, bullet_color=WHITE, font_size=12)
    add_bullet_box(slide, fitness,  Inches(4.55), Inches(1.55), Inches(4.2), Inches(5.7),
                   title="Fitness & Training", bg_color=LIGHT_BLUE,
                   title_color=ICE_BLUE, bullet_color=DARK_GREY, font_size=12)

    add_image_safe(slide, "skills", Inches(8.9), Inches(1.6), Inches(2.2), Inches(2.8))
    add_bullet_box(slide, nutrition, Inches(8.9), Inches(4.55), Inches(4.2), Inches(2.7),
                   title="Nutrition", bg_color=ICE_BLUE,
                   title_color=GOLD, bullet_color=WHITE, font_size=11)


def slide_05_competition(prs):
    """Slide 5 – Competition Process."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "The Competition Process",
                  "How Olympic Ice Hockey is played and structured")

    format_bullets = [
        "12 men's teams & 8 women's teams qualify for the Olympics.",
        "Teams are divided into groups for a preliminary round-robin stage.",
        "Top teams advance to a single-elimination playoff bracket.",
        "Quarterfinals → Semifinals → Bronze Medal Game → Gold Medal Game.",
        "Games are played on a 200 ft × 85 ft (61 m × 26 m) international rink.",
        "International rink is slightly wider than the NHL standard rink.",
        "IIHF (International Ice Hockey Federation) officiates all games.",
        "Two referees and two linesmen on the ice for each game.",
        "Video review (Coach's Challenge) available for disputed goals.",
    ]
    add_bullet_box(slide, format_bullets,
                   Inches(0.3), Inches(1.55), Inches(7.8), Inches(5.7),
                   title="Tournament Format & Game Structure",
                   bg_color=ICE_BLUE, title_color=GOLD, bullet_color=WHITE, font_size=13)

    add_image_safe(slide, "competition", Inches(8.3), Inches(1.6), Inches(4.7), Inches(3.0))

    key_phases = [
        "Face-off: puck drop starts each period & after stoppages",
        "Power Play: 5-on-4 advantage when opponent penalized",
        "Penalty Kill: defending a power play situation",
        "Line Changes: players substitute on the fly (no stoppage needed)",
        "Icing & Offside: automatic stoppage, face-off in defensive zone",
    ]
    add_bullet_box(slide, key_phases,
                   Inches(8.3), Inches(4.8), Inches(4.7), Inches(2.5),
                   title="Key In-Game Situations",
                   bg_color=LIGHT_BLUE, title_color=ICE_BLUE,
                   bullet_color=DARK_GREY, font_size=12)


def slide_06_video(prs):
    """Slide 6 – Video Clip."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, ICE_BLUE)
    add_title_bar(slide, "Watch Olympic Ice Hockey in Action",
                  "Highlight reel — see the speed, skill, and passion up close")

    # Big video link box
    add_rect(slide, Inches(1.5), Inches(1.7), Inches(10.3), Inches(4.0), DARK_GREY)
    add_text_box(slide, "▶  CLICK TO WATCH VIDEO",
                 Inches(1.5), Inches(2.8), Inches(10.3), Inches(0.7),
                 font_size=28, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    add_text_box(slide,
                 "Olympic Ice Hockey – Best Goals & Highlights (Beijing 2022)\n"
                 "https://www.youtube.com/watch?v=7oGfWJTFBF8\n"
                 "(Duration: approx. 2 min 45 sec)",
                 Inches(1.6), Inches(3.6), Inches(10.1), Inches(1.8),
                 font_size=16, bold=False, color=WHITE, align=PP_ALIGN.CENTER)

    add_text_box(slide,
                 "Backup video (official Olympic channel):\n"
                 "https://www.youtube.com/@Olympics  →  search \"Ice Hockey Highlights\"",
                 Inches(1.5), Inches(5.85), Inches(10.3), Inches(0.9),
                 font_size=13, bold=False, color=LIGHT_BLUE, align=PP_ALIGN.CENTER)

    add_image_safe(slide, "cover", Inches(0.2), Inches(2.5), Inches(1.1), Inches(1.8))
    add_image_safe(slide, "cover", Inches(12.0), Inches(2.5), Inches(1.1), Inches(1.8))


def slide_07_olympian(prs):
    """Slide 7 – Current Olympian: Connor McDavid."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "Spotlight Olympian: Connor McDavid",
                  "Team Canada | Center | 2026 Milan–Cortina Winter Olympics")

    bio = [
        "Born: January 13, 1997, in Richmond Hill, Ontario, Canada",
        "Position: Center  |  Shoots: Left  |  Height: 6'1\"  |  Weight: 193 lbs",
        "NHL Team: Edmonton Oilers (Captain since age 19 — youngest in NHL history)",
        "First Overall NHL Draft Pick (2015 NHL Entry Draft)",
        "Grew up playing hockey in Richmond Hill; turned pro at age 18",
        "Known as the fastest and most skilled player of his generation",
    ]
    achievements = [
        "6× Hart Trophy Winner (MVP) – 2017, 2021, 2022, 2023, 2024, 2025",
        "6× Art Ross Trophy (scoring title) – NHL's top point scorer",
        "Ted Lindsay Award (players' choice MVP) – multiple times",
        "NHL All-Star Game – selected every eligible season",
        "2025 Stanley Cup Champion with the Edmonton Oilers",
        "Team Canada – multiple World Championship gold medals",
        "2022 Beijing Winter Olympics – Team Canada (did not take gold but was standout player)",
        "2026 Milan–Cortina – expected to captain Team Canada",
    ]

    add_bullet_box(slide, bio,
                   Inches(0.3), Inches(1.55), Inches(6.5), Inches(2.85),
                   title="Background & Career", bg_color=ICE_BLUE,
                   title_color=GOLD, bullet_color=WHITE, font_size=12)
    add_bullet_box(slide, achievements,
                   Inches(0.3), Inches(4.55), Inches(6.5), Inches(2.75),
                   title="Major Achievements", bg_color=LIGHT_BLUE,
                   title_color=ICE_BLUE, bullet_color=DARK_GREY, font_size=12)

    add_image_safe(slide, "olympian", Inches(7.0), Inches(1.6), Inches(3.0), Inches(4.0))

    quote_box = [
        '"Every day I come to the rink, I want to be',
        ' the best player on the ice."',
        "                              — Connor McDavid",
    ]
    add_bullet_box(slide, quote_box,
                   Inches(7.0), Inches(5.8), Inches(6.0), Inches(1.5),
                   bg_color=GOLD, bullet_color=DARK_GREY, font_size=13)

    add_rect(slide, Inches(10.2), Inches(1.6), Inches(2.9), Inches(4.0), LIGHT_BLUE)
    stats = "2024-25 NHL Stats\n\nGP:  82\nG:   62\nA:   87\nPTS: 149\n+/-: +48\n\nCareer Points: 1,100+"
    add_text_box(slide, stats,
                 Inches(10.3), Inches(1.7), Inches(2.7), Inches(3.8),
                 font_size=13, bold=False, color=DARK_GREY, wrap=True)


def slide_08_quiz(prs):
    """Slide 8 – Final Exam Quiz Questions (4 multiple choice)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, ICE_BLUE)
    add_title_bar(slide, "Final Exam Review — Test Your Knowledge!",
                  "Choose the best answer for each question")

    questions = [
        {
            "q": "1.  In what year did men's ice hockey first appear at the Olympic Games?",
            "choices": ["A) 1896   B) 1900   C) 1920 ✓   D) 1936"],
            "note": "Answer: C — 1920 Antwerp Summer Olympics"
        },
        {
            "q": "2.  How many periods are in a regulation Olympic ice hockey game?",
            "choices": ["A) 2   B) 3 ✓   C) 4   D) 5"],
            "note": "Answer: B — 3 periods of 20 minutes each"
        },
        {
            "q": "3.  Which rule means a player cannot enter the offensive zone before the puck?",
            "choices": ["A) Icing   B) Hooking   C) High-Sticking   D) Offside ✓"],
            "note": "Answer: D — Offside"
        },
        {
            "q": "4.  Connor McDavid was selected _____ overall in the 2015 NHL Draft.",
            "choices": ["A) 2nd   B) 5th   C) 1st ✓   D) 10th"],
            "note": "Answer: C — 1st overall by the Edmonton Oilers"
        },
    ]

    colors = [LIGHT_BLUE, WHITE, LIGHT_BLUE, WHITE]
    text_c = [DARK_GREY, DARK_GREY, DARK_GREY, DARK_GREY]
    xs = [Inches(0.3), Inches(6.7), Inches(0.3), Inches(6.7)]
    ys = [Inches(1.6), Inches(1.6), Inches(4.55), Inches(4.55)]

    for i, (q_data, x, y, bg, tc) in enumerate(
            zip(questions, xs, ys, colors, text_c)):
        add_rect(slide, x, y, Inches(6.2), Inches(2.75), bg)
        add_text_box(slide, q_data["q"],
                     x + Inches(0.15), y + Inches(0.1),
                     Inches(5.9), Inches(0.9),
                     font_size=13, bold=True, color=ICE_BLUE, wrap=True)
        add_text_box(slide, q_data["choices"][0],
                     x + Inches(0.15), y + Inches(1.05),
                     Inches(5.9), Inches(0.6),
                     font_size=13, color=tc, wrap=True)
        add_text_box(slide, q_data["note"],
                     x + Inches(0.15), y + Inches(1.75),
                     Inches(5.9), Inches(0.55),
                     font_size=11, bold=True, color=RED, wrap=True)

    add_image_safe(slide, "quiz", Inches(5.9), Inches(6.4), Inches(1.5), Inches(0.9))


def slide_09_references(prs):
    """Slide 9 – References."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, WHITE)
    add_title_bar(slide, "References",
                  "Sources used for information in this presentation")

    refs = [
        "International Ice Hockey Federation (IIHF) — Official rules, history, statistics:\n    https://www.iihf.com",
        "International Olympic Committee — Olympic ice hockey history and results:\n    https://olympics.com/en/sports/ice-hockey/",
        "National Hockey League (NHL) — Player stats, bios, game rules:\n    https://www.nhl.com",
        "Connor McDavid — Official NHL player profile:\n    https://www.nhl.com/player/connor-mcdavid-8478402",
        "Olympic.org — Milan Cortina 2026 Winter Olympics:\n    https://olympics.com/en/olympic-games/milan-cortina-2026",
        "Wikipedia — Ice hockey at the Olympic Games:\n    https://en.wikipedia.org/wiki/Ice_hockey_at_the_Olympic_Games",
        "Wikimedia Commons — Images used throughout this presentation:\n    https://commons.wikimedia.org",
        "Hockey Canada — Athlete development, nutrition, and training guidelines:\n    https://www.hockeycanada.ca",
        "YouTube / Olympic Channel — Video highlights:\n    https://www.youtube.com/@Olympics",
    ]

    add_bullet_box(slide, refs,
                   Inches(0.3), Inches(1.55), Inches(8.8), Inches(5.7),
                   bg_color=LIGHT_BLUE, bullet_color=DARK_GREY, font_size=12)

    add_image_safe(slide, "refs", Inches(9.3), Inches(2.0), Inches(3.7), Inches(2.2))

    add_rect(slide, Inches(9.3), Inches(4.5), Inches(3.7), Inches(2.8), ICE_BLUE)
    add_text_box(slide,
                 "All images sourced from\nWikimedia Commons\n(Public Domain / CC License)\n\n"
                 "Video content from the\nOfficial Olympic YouTube Channel",
                 Inches(9.4), Inches(4.6), Inches(3.5), Inches(2.6),
                 font_size=13, color=WHITE, wrap=True)


# ════════════════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    print("Building slides...")
    builders = [
        ("Slide 1 – Cover",           slide_01_cover),
        ("Slide 2 – History",         slide_02_history),
        ("Slide 3 – Rules & Equipment",slide_03_rules),
        ("Slide 4 – Skills",          slide_04_skills),
        ("Slide 5 – Competition",     slide_05_competition),
        ("Slide 6 – Video",           slide_06_video),
        ("Slide 7 – Olympian",        slide_07_olympian),
        ("Slide 8 – Quiz",            slide_08_quiz),
        ("Slide 9 – References",      slide_09_references),
    ]
    for label, fn in builders:
        print(f"  {label}")
        fn(prs)

    out_path = os.path.join(os.path.dirname(__file__), "Olympic_Hockey_Presentation.pptx")
    prs.save(out_path)
    print(f"\nPresentation saved to:\n  {out_path}")
    print("\nTo use as Google Slides:")
    print("  1. Go to https://slides.google.com")
    print("  2. Click 'Blank presentation' → File → Import slides")
    print("     OR open drive.google.com → New → File Upload → open the .pptx")
    print("  3. Google Slides will auto-convert the file.")


if __name__ == "__main__":
    main()
