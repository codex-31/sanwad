#!/usr/bin/env python3
"""Build Sanwad (PS 26173, ISRO SAC) SIH 2026 deck — 7 slides, 16:9."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------- palette / constants ----------
NAVY   = RGBColor(0x0A, 0x1E, 0x3C)
NAVY2  = RGBColor(0x12, 0x2C, 0x50)
ORANGE = RGBColor(0xF4, 0x7B, 0x20)
INK    = RGBColor(0x1B, 0x2A, 0x44)
GRAY   = RGBColor(0x5C, 0x6B, 0x82)
LIGHT  = RGBColor(0xF3, 0xF7, 0xFB)
BORDER = RGBColor(0xD9, 0xE3, 0xEE)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
FONT   = "Calibri"

SW, SH = 13.333, 7.5
M = 0.6  # side margin

prs = Presentation()
prs.slide_width  = Inches(SW)
prs.slide_height = Inches(SH)
BLANK = prs.slide_layouts[6]

# ---------- helpers ----------
def slide_new(bg=WHITE):
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = bg
    r.line.fill.background(); r.shadow.inherit = False
    return s

def rect(s, x, y, w, h, fill=None, line=None, line_w=1.0, round_=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE):
    shp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if round_ is not None:
        try: shp.adjustments[0] = round_
        except Exception: pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp

def txt(s, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: list of dicts {runs:[(text,{size,bold,color,font,italic,spc})], align, sb, sa, ls}"""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, spec in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if spec.get("sb") is not None: p.space_before = Pt(spec["sb"])
        if spec.get("sa") is not None: p.space_after = Pt(spec["sa"])
        p.line_spacing = spec.get("ls", 1.06)
        for t, st in spec["runs"]:
            r = p.add_run(); r.text = t
            f = r.font
            f.size = Pt(st.get("size", 12))
            f.bold = st.get("bold", False)
            f.italic = st.get("italic", False)
            f.color.rgb = st.get("color", INK)
            f.name = st.get("font", FONT)
            if st.get("spc"):
                r._r.get_or_add_rPr().set("spc", str(st["spc"]))
    return tb

def bullets(s, x, y, w, h, items, size=12, gap=8, lead_color=NAVY, body_color=INK, ls=1.12):
    """items: list of (lead, body) or (None, body)."""
    paras = []
    for i, (lead, body) in enumerate(items):
        runs = [("•  ", {"size": size, "bold": True, "color": ORANGE})]
        if lead:
            runs.append((lead + " ", {"size": size, "bold": True, "color": lead_color}))
        runs.append((body, {"size": size, "color": body_color}))
        paras.append({"runs": runs, "sb": 0 if i == 0 else gap, "ls": ls})
    return txt(s, x, y, w, h, paras)

def header(s, kicker, title, num, title_size=26):
    txt(s, M, 0.34, SW - 2 * M, 0.3, [
        {"runs": [(kicker, {"size": 10.5, "bold": True, "color": ORANGE, "spc": 300})]}])
    txt(s, M, 0.62, SW - 2 * M, 0.75, [
        {"runs": [(title, {"size": title_size, "bold": True, "color": NAVY})]}])
    rect(s, M, 1.32, 1.15, 0.045, fill=ORANGE, shape=MSO_SHAPE.RECTANGLE)
    # footer
    txt(s, M, 7.12, 8, 0.25, [
        {"runs": [("Sanwad  ·  PS 26173  ·  ISRO SAC  ·  SIH 2026", {"size": 9, "color": GRAY})]}])
    txt(s, SW - M - 1, 7.12, 1, 0.25, [
        {"runs": [(num, {"size": 9, "bold": True, "color": GRAY})], "align": PP_ALIGN.RIGHT}])

def card(s, x, y, w, h, fill=LIGHT, line=BORDER, radius=0.055):
    return rect(s, x, y, w, h, fill=fill, line=line, line_w=1.0, round_=radius)

def card_title(s, x, y, w, text, color=NAVY, size=13, spc=200):
    txt(s, x, y, w, 0.35, [{"runs": [(text, {"size": size, "bold": True, "color": color, "spc": spc})]}])

# =====================================================================
# SLIDE 1 — TITLE
# =====================================================================
s = slide_new(NAVY)
rect(s, 0, 0, SW, 0.14, fill=ORANGE, shape=MSO_SHAPE.RECTANGLE)
txt(s, M, 1.02, SW - 2 * M, 0.35, [
    {"runs": [("SMART INDIA HACKATHON 2026  ·  PROBLEM STATEMENT 26173  ·  ISRO SPACE APPLICATIONS CENTRE",
               {"size": 11.5, "bold": True, "color": ORANGE, "spc": 250})]}])
# title (Devanagari gets Nirmala UI; falls back gracefully)
txt(s, M, 1.45, SW - 2 * M, 1.5, [
    {"runs": [
        ("Sanwad ", {"size": 60, "bold": True, "color": WHITE}),
        ("संवाद", {"size": 40, "bold": True, "color": ORANGE, "font": "Nirmala UI"}),
    ]}])
txt(s, M, 3.05, SW - 2 * M, 0.5, [
    {"runs": [("Voice in.  Text across.  Voice out.", {"size": 21, "bold": True, "color": WHITE})]}])
txt(s, M, 3.62, 10.8, 1.1, [
    {"runs": [("A fully offline, open-source emergency walkie-talkie for Android. Local STT + TTS in "
               "10 Indian languages — built for the links where audio simply cannot fit.",
               {"size": 13.5, "color": RGBColor(0xC9, 0xD5, 0xE6)})], "ls": 1.25}])
# chips
chips = ["100% ON-DEVICE", "10 INDIAN LANGUAGES", "OPEN SOURCE ONLY"]
cx = M
for c in chips:
    wch = 1.95 + 0.011 * len(c)
    card(s, cx, 5.05, wch, 0.5, fill=NAVY2, line=RGBColor(0x2E, 0x47, 0x6E), radius=0.5)
    txt(s, cx, 5.05, wch, 0.5, [{"runs": [(c, {"size": 10.5, "bold": True, "color": WHITE, "spc": 150})],
                                 "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
    cx += wch + 0.25
txt(s, M, 6.35, SW - 2 * M, 0.4, [
    {"runs": [
        ("Team:  ______________________________", {"size": 12.5, "color": RGBColor(0xC9, 0xD5, 0xE6)}),
        ("          Team ID:  ", {"size": 12.5, "color": RGBColor(0xC9, 0xD5, 0xE6)}),
        ("______________", {"size": 12.5, "bold": True, "color": ORANGE}),
    ]}])

# =====================================================================
# SLIDE 2 — THE PROBLEM
# =====================================================================
s = slide_new()
header(s, "THE PROBLEM", "When the network dies, voice is the last link", "02")
LX, LW = M, 6.95
bullets(s, LX, 1.62, LW, 3.6, [
    ("Disasters take the network first.",
     "Floods, cyclones, quakes and coastal operations cut cellular exactly when communication matters most."),
    ("Emergency links are narrowband.",
     "Satellite / HF–VHF channels carry a few hundred bits per second — not kilobits."),
    ("Voice is 8–16 kbps, continuous.",
     "It physically cannot fit a few-hundred-bps link, however compressed."),
    ("SMS fits — but context doesn't.",
     "Injuries, location, boat condition: in a crisis, people must speak, not type."),
    ("Panic, noise, non-literate users.",
     "Speaking is the most inclusive, natural interface — if the link can carry it."),
], size=12)
# real-world example card
RX, RW = 7.85, SW - M - 7.85
card(s, RX, 1.62, RW, 3.55)
rect(s, RX, 1.62, 0.07, 3.55, fill=ORANGE, shape=MSO_SHAPE.RECTANGLE)
card_title(s, RX + 0.32, 1.88, RW - 0.6, "REAL-WORLD EXAMPLE", color=ORANGE, size=11.5)
txt(s, RX + 0.32, 2.32, RW - 0.62, 2.7, [
    {"runs": [("A fisher off the Konkan coast: one crewman injured, boat taking on water. "
               "Cellular is dead; a voice call is impossible over the low-rate link. "
               "SMS can't say “third man — arm bleed — leaking fast” — "
               "and nobody can answer by voice.",
               {"size": 12.5, "color": INK})], "ls": 1.28}])
# hidden-trap bar
rect(s, M, 5.55, SW - 2 * M, 1.05, fill=NAVY, round_=0.12)
txt(s, M + 0.35, 5.55, SW - 2 * M - 0.7, 1.05, [
    {"runs": [
        ("The hidden trap:  ", {"size": 13, "bold": True, "color": ORANGE}),
        ("the phone “always works” — until it is the only thing that doesn't. "
         "The last mile of crisis communication is a link too thin for audio.",
         {"size": 13, "italic": True, "color": WHITE}),
    ], "ls": 1.2}], anchor=MSO_ANCHOR.MIDDLE)

# =====================================================================
# SLIDE 3 — WHY DIFFERENT
# =====================================================================
s = slide_new()
header(s, "WHY SANWAD IS DIFFERENT", "50–100× less data than audio — the same voice experience", "03")
txt(s, M, 1.5, SW - 2 * M, 0.3, [
    {"runs": [("Per 5-second utterance — same units, same window:",
               {"size": 11, "italic": True, "color": GRAY})]}])
stats = [("~100 B", "Sanwad: one sentence as text"),
         ("40–80 KB", "raw voice at 8–16 kbps"),
         ("50–100×", "less data crosses the link")]
gw = (SW - 2 * M - 0.4) / 3
for i, (big, small) in enumerate(stats):
    x = M + i * (gw + 0.2)
    card(s, x, 1.88, gw, 1.5)
    txt(s, x, 2.08, gw, 0.65, [{"runs": [(big, {"size": 30, "bold": True, "color": NAVY})],
                                "align": PP_ALIGN.CENTER}])
    txt(s, x + 0.15, 2.82, gw - 0.3, 0.45, [{"runs": [(small, {"size": 11, "color": GRAY})],
                                             "align": PP_ALIGN.CENTER}])
# core framing box
rect(s, M, 3.72, SW - 2 * M, 1.95, fill=NAVY, round_=0.07)
txt(s, M + 0.4, 3.98, SW - 2 * M - 0.8, 1.5, [
    {"runs": [("Text-first, voice-at-both-ends.  ", {"size": 14, "bold": True, "color": ORANGE})]},
    {"runs": [("The phone speaks at both ends — STT turns speech into text, TTS turns it back — "
               "but only ~100 B of text crosses the link. That one decision is why the same app runs over "
               "Wi-Fi Direct / Bluetooth today, and ports to LoRa, NavIC-ecosystem messaging and HF/VHF "
               "packet links later — no re-architecture.",
               {"size": 13, "color": WHITE})], "sb": 6, "ls": 1.3}])
txt(s, M, 6.0, SW - 2 * M, 0.75, [
    {"runs": [
        ("vs. existing walkie-talkie apps:  ", {"size": 12, "bold": True, "color": NAVY}),
        ("cellular-based (die with the network) or VHF radio hardware (new equipment). "
         "Sanwad is the phone you already carry.", {"size": 12, "color": GRAY}),
    ], "ls": 1.25}])

# =====================================================================
# SLIDE 4 — ARCHITECTURE
# =====================================================================
s = slide_new()
header(s, "TECHNICAL APPROACH", "How a word becomes a voice on the other phone", "04")

def flow_box(x, y, w, h, l1, l2, fill=WHITE, l1c=NAVY, l2c=GRAY):
    card(s, x, y, w, h, fill=fill, line=BORDER, radius=0.09)
    txt(s, x + 0.12, y, w - 0.24, h, [
        {"runs": [(l1, {"size": 11.5, "bold": True, "color": l1c})], "align": PP_ALIGN.CENTER},
        {"runs": [(l2, {"size": 9.5, "color": l2c})], "align": PP_ALIGN.CENTER, "sb": 2}],
        anchor=MSO_ANCHOR.MIDDLE)

def chev(x, y):
    c = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(0.3), Inches(0.34))
    c.fill.solid(); c.fill.fore_color.rgb = ORANGE; c.line.fill.background(); c.shadow.inherit = False

txt(s, M, 1.5, 6, 0.3, [{"runs": [("PHONE A  ·  SPEAK → TEXT", {"size": 10.5, "bold": True, "color": ORANGE, "spc": 200})]}])
yA = 1.88; hA = 1.25
flow_box(M,       yA, 2.35, hA, "Mic + VAD", "pause & stoppage detection")
chev(3.02, yA + hA/2 - 0.17)
flow_box(3.40,    yA, 3.30, hA, "On-device STT", "IndicConformer · 10 languages · int8 TFLite")
chev(6.78, yA + hA/2 - 0.17)
flow_box(7.16,    yA, 2.60, hA, "Sentence JSON", "≈ 100 bytes per sentence")
# transport bar
rect(s, M, 3.42, SW - 2 * M, 0.78, fill=NAVY, round_=0.16)
txt(s, M + 0.35, 3.42, 6.5, 0.78, [
    {"runs": [("Wi-Fi Direct / Bluetooth LE", {"size": 13, "bold": True, "color": WHITE})]}],
    anchor=MSO_ANCHOR.MIDDLE)
rect(s, 8.1, 3.58, 4.28, 0.46, fill=ORANGE, round_=0.5)
txt(s, 8.1, 3.58, 4.28, 0.46, [
    {"runs": [("TEXT ONLY — AUDIO NEVER CROSSES", {"size": 9.5, "bold": True, "color": WHITE, "spc": 100})],
      "align": PP_ALIGN.CENTER}], anchor=MSO_ANCHOR.MIDDLE)
# receiver lane
txt(s, M, 4.52, 6, 0.3, [{"runs": [("PHONE B  ·  TEXT → SPEAK", {"size": 10.5, "bold": True, "color": ORANGE, "spc": 200})]}])
yB = 4.90; hB = 1.25
flow_box(M,       yB, 2.35, hB, "Local TTS", "per-language model, on-device")
chev(3.02, yB + hB/2 - 0.17)
flow_box(3.40,    yB, 3.30, hB, "Normal message", "played back as a voice note")
chev(6.78, yB + hB/2 - 0.17)
flow_box(7.16,    yB, 3.90, hB, "Alert message", "highest volume · non-interruptible · bypasses mute")
# bottom strip
txt(s, M, 6.42, SW - 2 * M, 0.5, [
    {"runs": [
        ("Push-to-talk walkie-talkie UI  ·  same app on both phones  ·  ",
         {"size": 11.5, "color": GRAY}),
        ("behaves like a normal phone when the mode is off", {"size": 11.5, "bold": True, "color": NAVY}),
    ], "align": PP_ALIGN.CENTER}])

# =====================================================================
# SLIDE 5 — TECH CHOICES
# =====================================================================
s = slide_new()
header(s, "TECHNOLOGY CHOICES", "What we build with — and what we rejected", "05")
cw = (SW - 2 * M - 0.25) / 2
y5 = 1.62; h5 = 3.95
# STT card
card(s, M, y5, cw, h5)
card_title(s, M + 0.3, y5 + 0.25, cw - 0.6, "STT — SPEECH TO TEXT")
bullets(s, M + 0.3, y5 + 0.72, cw - 0.6, h5 - 1.0, [
    ("Primary: IndicConformer (AI4Bharat).", "TFLite int8, fully on-device, 10 languages."),
    ("Fallback: Whisper.cpp.", "Broader accent robustness at a larger footprint."),
    ("Evaluated, not selected: Vosk.", "Indic language coverage below this brief's bar."),
    ("Design rule:", "STT fires on detected pause; emits one compact sentence at a time."),
], size=11.5, gap=9)
# TTS card
card(s, M + cw + 0.25, y5, cw, h5)
card_title(s, M + cw + 0.55, y5 + 0.25, cw - 0.6, "TTS — TEXT TO SPEECH")
bullets(s, M + cw + 0.55, y5 + 0.72, cw - 0.6, h5 - 1.0, [
    ("Baseline, all 10 languages: espeak-ng.", "Fully offline, tiny, fast, intelligible."),
    ("Neural where supported: Piper (en / hi).", "Better naturalness for high-frequency alert phrases."),
    ("Priority: speed + intelligibility", "over studio fidelity — this is emergency comms."),
    ("Alerts:", "highest volume, non-interruptible playback, bypasses system mute."),
], size=11.5, gap=9)
# chips
chips5 = [("100% OFFLINE", "no cloud, no cellular"),
          ("OPEN SOURCE ONLY", "Apache-2.0 / MIT · license audit in progress"),
          ("2 GB RAM TARGET", "₹8–10k · Cortex-A53-class devices")]
x = M
for big, small in chips5:
    wch = (SW - 2 * M - 0.5) / 3
    card(s, x, 5.85, wch, 0.85, fill=NAVY, line=None, radius=0.12)
    txt(s, x + 0.25, 5.85, wch - 0.5, 0.85, [
        {"runs": [(big, {"size": 10.5, "bold": True, "color": ORANGE, "spc": 100})]},
        {"runs": [(small, {"size": 10, "color": WHITE})], "sb": 3}], anchor=MSO_ANCHOR.MIDDLE)
    x += wch + 0.25

# =====================================================================
# SLIDE 6 — METRICS
# =====================================================================
s = slide_new()
header(s, "AGAINST THE EVALUATION RUBRIC", "Efficiency · Accuracy · Latency — targets we can prove", "06")
colw = (SW - 2 * M - 0.4) / 3
y6 = 1.58; h6 = 4.55
cols = [
    ("EFFICIENCY", "20%", [
        ("STT model", "< 300 MB int8, on-device"),
        ("App footprint", "< 500 MB installed"),
        ("Memory", "2 GB RAM target"),
        ("Idle listening", "< 5% CPU (VAD-gated decode)"),
        ("Operating cost", "₹0 recurring cloud / cellular"),
    ]),
    ("ACCURACY", "40%", [
        ("STT WER", "target on in-domain test set, all 10 languages"),
        ("TTS", "human legibility + flow panel"),
        ("Safety net", "text preview before playback — errors visible, fixable"),
        ("Tuning loop", "fine-tune / swap model per language if WER misses"),
    ]),
    ("LATENCY", "20%", None),
]
for i, (t, pct, items) in enumerate(cols):
    x = M + i * (colw + 0.2)
    card(s, x, y6, colw, h6)
    card_title(s, x + 0.3, y6 + 0.24, colw - 0.6, t)
    rect(s, x + colw - 1.05, y6 + 0.22, 0.78, 0.4, fill=ORANGE, round_=0.5)
    txt(s, x + colw - 1.05, y6 + 0.22, 0.78, 0.4, [
        {"runs": [(pct, {"size": 11, "bold": True, "color": WHITE})], "align": PP_ALIGN.CENTER}],
        anchor=MSO_ANCHOR.MIDDLE)
    if items:
        paras = []
        for j, (k, v) in enumerate(items):
            paras.append({"runs": [("•  ", {"size": 11, "bold": True, "color": ORANGE}),
                                   (k + "  ", {"size": 11, "bold": True, "color": NAVY}),
                                   (v, {"size": 11, "color": INK})],
                          "sb": 0 if j == 0 else 8, "ls": 1.18})
        txt(s, x + 0.3, y6 + 0.85, colw - 0.6, h6 - 1.1, paras)
# latency column: budget table
x3 = M + 2 * (colw + 0.2)
rows = [("STT — pause detect + decode", "1.5–2.5 s"),
        ("Transport — Wi-Fi Direct / BLE", "< 0.5 s"),
        ("TTS — synthesize + start", "< 1.5 s"),
        ("Mouth-to-ear total", "≤ 4–5 s")]
ry = y6 + 0.95
for k, (label, val) in enumerate(rows):
    if k > 0:
        rect(s, x3 + 0.3, ry - 0.075, colw - 0.6, 0.012, fill=BORDER, shape=MSO_SHAPE.RECTANGLE)
    hl = (k == len(rows) - 1)
    txt(s, x3 + 0.3, ry, colw - 1.35, 0.55, [
        {"runs": [(label, {"size": 11, "bold": hl, "color": NAVY if hl else INK})], "ls": 1.1}])
    txt(s, x3 + colw - 1.3, ry, 1.0, 0.55, [
        {"runs": [(val, {"size": 12, "bold": True, "color": ORANGE if hl else NAVY})],
          "align": PP_ALIGN.RIGHT}])
    ry += 0.62
txt(s, x3 + 0.3, ry + 0.02, colw - 0.6, 0.5, [
    {"runs": [("RTF target < 0.5  ·  measured on target device", {"size": 10, "italic": True, "color": GRAY})]}])
# note
txt(s, M, 6.35, SW - 2 * M, 0.55, [
    {"runs": [
        ("Honest note:  ", {"size": 11, "bold": True, "color": ORANGE}),
        ("these are design targets, not measurements. Benchmarking is planned on a ₹8–10k device; "
         "we will present measured numbers at demo time.", {"size": 11, "color": GRAY}),
    ], "ls": 1.2}])

# =====================================================================
# SLIDE 7 — ROADMAP & ADOPTION
# =====================================================================
s = slide_new()
header(s, "ROADMAP & ADOPTION", "From a two-phone demo to a district-scale service", "07")
ph = [
    ("PHASE 1 · NOW (SIH)",
     "Complete Android app. Live walkie-talkie demo: two phones over Wi-Fi Direct / Bluetooth, "
     "STT→TTS in 10 languages, with an on-screen packet log proving only text crosses the link."),
    ("PHASE 2 · FIELD",
     "LoRa bridge for mesh relay. Low-rate messaging study (NavIC-ecosystem / HF–VHF packet). "
     "Pilots with NDRF and state maritime boards."),
    ("PHASE 3 · SCALE",
     "B2G distribution with district administrations (DPI). Multilingual alert library. "
     "Per-language model updates."),
]
for i, (t, b) in enumerate(ph):
    x = M + i * (colw + 0.2)
    card(s, x, 1.58, colw, 2.55)
    rect(s, x, 1.58, colw, 0.09, fill=ORANGE if i == 0 else BORDER, shape=MSO_SHAPE.RECTANGLE)
    card_title(s, x + 0.3, 1.85, colw - 0.6, t, size=12, spc=150)
    txt(s, x + 0.3, 2.3, colw - 0.6, 1.7, [
        {"runs": [(b, {"size": 11, "color": INK})], "ls": 1.25}])
# adoption path
card(s, M, 4.42, SW - 2 * M, 0.72, fill=LIGHT, line=BORDER)
txt(s, M + 0.35, 4.42, SW - 2 * M - 0.7, 0.72, [
    {"runs": [
        ("Target adoption path (hypothesis to test, not assumption):  ", {"size": 11.5, "bold": True, "color": NAVY}),
        ("NDRF  →  state maritime boards  →  district administrations (DPI).  Pilots in negotiation.",
         {"size": 11.5, "color": INK}),
    ]}], anchor=MSO_ANCHOR.MIDDLE)
# ask
rect(s, M, 5.38, SW - 2 * M, 0.78, fill=WHITE, line=ORANGE, line_w=1.5, round_=0.08)
txt(s, M + 0.35, 5.38, SW - 2 * M - 0.7, 0.78, [
    {"runs": [
        ("We ask ISRO:  ", {"size": 11.5, "bold": True, "color": ORANGE}),
        ("guidance on low-rate messaging specs (bit budget, update interval) and support for a field pilot.",
         {"size": 11.5, "color": INK}),
    ]}], anchor=MSO_ANCHOR.MIDDLE)
# closing
rect(s, M, 6.42, SW - 2 * M, 0.55, fill=NAVY, round_=0.2)
txt(s, M, 6.42, SW - 2 * M, 0.55, [
    {"runs": [("Zero recurring cloud or cellular cost.  Words travel even when data can't.",
               {"size": 13, "bold": True, "color": WHITE})], "align": PP_ALIGN.CENTER}],
    anchor=MSO_ANCHOR.MIDDLE)

prs.save("/home/user/Sanwad_PS26173_SIHS2026_Deck.pptx")
print("saved", len(prs.slides._sldIdLst), "slides")
