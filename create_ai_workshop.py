from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Bright Palette ────────────────────────────────────────────────────────────
C_BG      = RGBColor(0xFF, 0xFF, 0xFF)   # white        – slide backgrounds
C_BG2     = RGBColor(0xF4, 0xF6, 0xFF)  # soft lavender – panel backgrounds
C_PURPLE  = RGBColor(0x6C, 0x35, 0xDE)  # vivid purple  – primary headers
C_BLUE    = RGBColor(0x22, 0x8B, 0xE6)  # bright blue   – accents / subheaders
C_TEAL    = RGBColor(0x00, 0xBF, 0xA5)  # teal green    – positive / correct
C_GREEN   = RGBColor(0x00, 0xC8, 0x53)  # lime green    – can-do / allowed
C_ORANGE  = RGBColor(0xFF, 0x6D, 0x00)  # vivid orange  – activities / warnings
C_PINK    = RGBColor(0xE9, 0x1E, 0x8C)  # hot pink      – caution / wrong
C_YELLOW  = RGBColor(0xFF, 0xD6, 0x00)  # bright yellow – quiz / emphasis
C_DARK    = RGBColor(0x1A, 0x1A, 0x2E)  # near-black    – body text on light bg
C_DGRAY   = RGBColor(0x44, 0x44, 0x66)  # dark grey     – secondary text
C_LGRAY   = RGBColor(0xE8, 0xEC, 0xF8)  # light grey    – panel fills
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]   # completely blank

# ─────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────────────────────────────────────

def add_slide(bg=None):
    sl = prs.slides.add_slide(BLANK)
    fill = sl.background.fill
    fill.solid()
    fill.fore_color.rgb = bg if bg else C_BG
    return sl


def txb(sl, text, x, y, w, h,
        size=24, bold=False, color=C_DARK, align=PP_ALIGN.LEFT,
        italic=False, wrap=True):
    box = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = box.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(size)
    run.font.bold   = bold
    run.font.color.rgb = color
    run.font.italic    = italic
    return run


def rect(sl, x, y, w, h, fill_color):
    shape = sl.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def rounded_rect(sl, x, y, w, h, fill_color):
    """Rounded rectangle (MSO_SHAPE freeform – use rounded_rectangle preset id=5)."""
    from pptx.util import Emu
    from pptx.oxml.ns import qn
    shape = sl.shapes.add_shape(
        5,  # rounded rectangle
        Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def header_bar(sl, title, subtitle=None,
               bar_color=C_PURPLE, title_color=C_WHITE, sub_color=C_LGRAY):
    rect(sl, 0, 0, 13.33, 1.45, bar_color)
    txb(sl, title, 0.35, 0.1, 12.3, 0.85,
        size=36, bold=True, color=title_color, align=PP_ALIGN.LEFT)
    if subtitle:
        txb(sl, subtitle, 0.4, 0.95, 12.3, 0.48,
            size=18, color=sub_color, italic=True)


def section_divider(section_num, section_title, subtitle="", accent=C_PURPLE):
    sl = add_slide(bg=accent)
    # large faded number watermark
    txb(sl, str(section_num), 8.5, 0.5, 4.5, 6.5,
        size=280, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
        align=PP_ALIGN.LEFT)
    # left accent bar
    rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
    txb(sl, f"SECTION {section_num}", 0.85, 2.1, 10, 0.55,
        size=18, bold=True, color=C_YELLOW, italic=True)
    txb(sl, section_title, 0.85, 2.75, 10.5, 1.5,
        size=46, bold=True, color=C_WHITE)
    if subtitle:
        txb(sl, subtitle, 0.85, 4.4, 10.5, 0.85,
            size=22, color=C_LGRAY, italic=True)
    return sl


def bullet_slide(title, bullets, subtitle=None,
                 bullet_color=C_DARK, icon="▸", size=21,
                 accent=C_PURPLE):
    sl = add_slide()
    header_bar(sl, title, subtitle, bar_color=accent)
    box = sl.shapes.add_textbox(Inches(0.6), Inches(1.6), Inches(12.1), Inches(5.6))
    tf  = box.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(9)
        run = p.add_run()
        run.text = f"{icon}  {b}" if icon else b
        run.font.size      = Pt(size)
        run.font.color.rgb = bullet_color
    return sl


def two_col_slide(title, left_title, left_items,
                  right_title, right_items,
                  left_col=C_PINK, right_col=C_TEAL,
                  subtitle=None, accent=C_PURPLE):
    sl = add_slide()
    header_bar(sl, title, subtitle, bar_color=accent)
    # left panel
    rect(sl, 0.3, 1.55, 6.15, 5.65, left_col)
    txb(sl, left_title, 0.5, 1.65, 5.75, 0.72,
        size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    box_l = sl.shapes.add_textbox(Inches(0.5), Inches(2.48), Inches(5.75), Inches(4.55))
    tf_l  = box_l.text_frame; tf_l.word_wrap = True
    first = True
    for item in left_items:
        if first: p = tf_l.paragraphs[0]; first = False
        else: p = tf_l.add_paragraph()
        p.space_before = Pt(8)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(19); run.font.color.rgb = C_WHITE
    # right panel
    rect(sl, 6.88, 1.55, 6.15, 5.65, right_col)
    txb(sl, right_title, 7.08, 1.65, 5.75, 0.72,
        size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    box_r = sl.shapes.add_textbox(Inches(7.08), Inches(2.48), Inches(5.75), Inches(4.55))
    tf_r  = box_r.text_frame; tf_r.word_wrap = True
    first = True
    for item in right_items:
        if first: p = tf_r.paragraphs[0]; first = False
        else: p = tf_r.add_paragraph()
        p.space_before = Pt(8)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(19); run.font.color.rgb = C_WHITE
    return sl


def activity_slide(title, instructions, time_str="5 min", color=C_ORANGE):
    sl = add_slide()
    rect(sl, 0, 0, 13.33, 1.45, color)
    txb(sl, f"✏️  Activity: {title}", 0.35, 0.1, 10.8, 0.82,
        size=32, bold=True, color=C_WHITE)
    txb(sl, f"⏱  {time_str}", 10.8, 0.15, 2.3, 0.65,
        size=22, bold=True, color=C_WHITE, align=PP_ALIGN.RIGHT)
    rect(sl, 0.4, 1.6, 12.53, 5.55, C_LGRAY)
    box = sl.shapes.add_textbox(Inches(0.75), Inches(1.8), Inches(12.0), Inches(5.15))
    tf  = box.text_frame; tf.word_wrap = True
    first = True
    for line in instructions:
        if first: p = tf.paragraphs[0]; first = False
        else: p = tf.add_paragraph()
        p.space_before = Pt(9)
        run = p.add_run(); run.text = line
        run.font.size = Pt(21)
        if line.startswith("→"):
            run.font.color.rgb = C_PURPLE; run.font.bold = True
        else:
            run.font.color.rgb = C_DARK
    return sl


def example_slide(title, prompt_text, output_text, subtitle=None, accent=C_PURPLE):
    sl = add_slide()
    header_bar(sl, title, subtitle, bar_color=accent)
    # prompt label strip
    rect(sl, 0.3, 1.55, 12.73, 0.42, C_BLUE)
    txb(sl, "PROMPT", 0.5, 1.58, 4, 0.36,
        size=14, bold=True, color=C_WHITE)
    rect(sl, 0.3, 1.97, 12.73, 2.05, C_LGRAY)
    box_p = sl.shapes.add_textbox(Inches(0.5), Inches(2.02), Inches(12.3), Inches(1.9))
    tf_p  = box_p.text_frame; tf_p.word_wrap = True
    run = tf_p.paragraphs[0].add_run()
    run.text = prompt_text
    run.font.size = Pt(19); run.font.color.rgb = C_PURPLE; run.font.italic = True
    # output label strip
    rect(sl, 0.3, 4.12, 12.73, 0.42, C_TEAL)
    txb(sl, "AI OUTPUT", 0.5, 4.15, 5, 0.36,
        size=14, bold=True, color=C_WHITE)
    rect(sl, 0.3, 4.54, 12.73, 2.65, C_LGRAY)
    box_o = sl.shapes.add_textbox(Inches(0.5), Inches(4.59), Inches(12.3), Inches(2.5))
    tf_o  = box_o.text_frame; tf_o.word_wrap = True
    run2 = tf_o.paragraphs[0].add_run()
    run2.text = output_text
    run2.font.size = Pt(18); run2.font.color.rgb = C_DGRAY
    return sl


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1  –  Title / Cover
# ═════════════════════════════════════════════════════════════════════════════
sl = add_slide(bg=C_PURPLE)
# diagonal colour band
rect(sl, 0, 0, 7.0, 7.5, C_BLUE)
rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
# title block
txb(sl, "AI as Your", 0.9, 1.0, 10, 1.15,
    size=56, bold=True, color=C_WHITE)
txb(sl, "Collaborator", 0.9, 2.2, 10, 1.15,
    size=56, bold=True, color=C_YELLOW)
txb(sl, "An Interactive Workshop for International College Students",
    0.9, 3.55, 9.5, 0.85, size=24, color=C_WHITE, italic=True)
rect(sl, 0.9, 4.55, 9.5, 0.07, C_YELLOW)
txb(sl, "Basics  ·  Mindset Shift  ·  Prompting Lab  ·  Evaluation  ·  Responsible Use",
    0.9, 4.75, 9.5, 0.6, size=17, color=C_LGRAY)
txb(sl, "Facilitator Guide  |  Workshop Version 1.0",
    0.9, 5.55, 7, 0.5, size=15, color=C_LGRAY, italic=True)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 0  –  OPENING QUIZ
# ═════════════════════════════════════════════════════════════════════════════
section_divider(0, "Opening Quiz",
                "Let's see what you already know — no grades, just fun!",
                accent=RGBColor(0xE9, 0x1E, 0x8C))

# --- Quiz Instructions ---
sl = add_slide()
header_bar(sl, "How the Quiz Works",
           subtitle="Scan the QR code — answer on your device!",
           bar_color=C_PINK, title_color=C_WHITE, sub_color=C_YELLOW)
rect(sl, 0.4, 1.58, 8.2, 5.6, C_LGRAY)
box = sl.shapes.add_textbox(Inches(0.7), Inches(1.78), Inches(7.8), Inches(5.2))
tf  = box.text_frame; tf.word_wrap = True
instructions = [
    "📋  12 quick questions across 4 categories.",
    "🤔  Answer honestly — this is NOT graded.",
    "🤫  Results are for your reflection only.",
    "📱  Use your phone or laptop to respond.",
    "⏱   About 1–2 minutes per question.",
    "🎯  Goal: understand where you start today.",
    "",
    "Categories:",
    "   AI Basics  |  Mindset  |  Prompting Clarity  |  Responsible Use",
]
first = True
for line in instructions:
    if first: p = tf.paragraphs[0]; first = False
    else: p = tf.add_paragraph()
    p.space_before = Pt(6)
    run = p.add_run(); run.text = line
    run.font.size = Pt(20)
    if line.startswith("   "):
        run.font.color.rgb = C_PURPLE; run.font.bold = True
    elif line.startswith("Categories"):
        run.font.color.rgb = C_PINK; run.font.bold = True
    else:
        run.font.color.rgb = C_DARK

# QR placeholder on the right
rect(sl, 9.0, 1.7, 3.9, 3.9, C_LGRAY)
rect(sl, 9.15, 1.85, 3.6, 3.6, C_WHITE)
# inner QR grid lines to suggest a QR code visually
for gx in range(3):
    for gy in range(3):
        if (gx == gy == 0) or (gx == 0 and gy == 2) or (gx == 2 and gy == 0):
            rect(sl, 9.25 + gx*1.1, 1.95 + gy*1.1, 0.9, 0.9, C_DARK)
        elif gx == gy == 1:
            rect(sl, 9.65, 2.35, 0.3, 0.3, C_DARK)
        else:
            rect(sl, 9.3 + gx*1.1, 2.0 + gy*1.1, 0.7, 0.7, C_LGRAY)
txb(sl, "SCAN ME", 9.0, 5.7, 3.9, 0.55,
    size=18, bold=True, color=C_PINK, align=PP_ALIGN.CENTER)
txb(sl, "Place your quiz link / Mentimeter /\nPoll Everywhere QR code here",
    8.9, 6.3, 4.1, 0.9, size=13, color=C_DGRAY, align=PP_ALIGN.CENTER, italic=True)

# --- Quiz Debrief ---
sl = add_slide()
header_bar(sl, "Quiz Debrief", subtitle="Let's talk about what we discovered",
           bar_color=C_TEAL, title_color=C_WHITE, sub_color=C_YELLOW)
rect(sl, 0.4, 1.58, 12.53, 5.6, C_LGRAY)
box = sl.shapes.add_textbox(Inches(0.7), Inches(1.78), Inches(12.0), Inches(5.2))
tf  = box.text_frame; tf.word_wrap = True
debrief = [
    "✅  Q1–Q3  |  AI Basics:  AI predicts language — it doesn't 'know' truth.",
    "✅  Q4–Q5  |  Mindset:  Specific, collaborative prompts get better results.",
    "✅  Q6–Q8  |  Prompting:  Role, format, and follow‑up questions improve AI output.",
    "✅  Q9–Q12 |  Responsible Use:  AI is a tool — how you use it is your choice.",
    "",
    "Key takeaway:",
    "→  The goal of this workshop is to move from passive user to active collaborator.",
    "→  Every strategy we practice today will make your AI interactions more powerful.",
]
first = True
for line in debrief:
    if first: p = tf.paragraphs[0]; first = False
    else: p = tf.add_paragraph()
    p.space_before = Pt(9)
    run = p.add_run(); run.text = line
    run.font.size = Pt(21)
    if line.startswith("→"):
        run.font.color.rgb = C_PURPLE; run.font.bold = True
    elif line.startswith("Key"):
        run.font.color.rgb = C_ORANGE; run.font.bold = True
    else:
        run.font.color.rgb = C_DARK

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1  –  BASICS OF AI
# ═════════════════════════════════════════════════════════════════════════════
section_divider(1, "Basics of AI",
                "What it is, what it can do, and where it falls short",
                accent=C_BLUE)

# What AI IS
sl = add_slide()
header_bar(sl, "What Is AI?", subtitle="A simple explanation for everyday use",
           bar_color=C_BLUE)
concepts = [
    ("🔍 Pattern Recognition", C_BLUE,
     "AI learns by finding patterns in enormous amounts of text, images, and data. "
     "It does not 'think' — it matches patterns."),
    ("💬 Language Generation", C_PURPLE,
     "AI chatbots predict the most likely next word — like autocomplete on a massive scale. "
     "Each word is chosen statistically."),
    ("🚫 No Understanding", C_PINK,
     "AI does not truly 'know' things. It has no beliefs, no memory between sessions "
     "(by default), and no lived experience."),
]
for i, (title, col, desc) in enumerate(concepts):
    bx = 0.3 + i * 4.35; by = 1.55
    rect(sl, bx, by, 4.1, 5.65, col)
    txb(sl, title, bx + 0.15, by + 0.2, 3.8, 0.75,
        size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    rect(sl, bx + 0.1, by + 1.05, 3.9, 0.06, C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.15), Inches(by + 1.2), Inches(3.8), Inches(4.1))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(18); run.font.color.rgb = C_WHITE

txb(sl, "🔍  Analogy: AI is like very well-read autocomplete — not a search engine, not a brain.",
    0.3, 7.05, 12.7, 0.42, size=15, color=C_BLUE, italic=True)

# What AI CAN DO
bullet_slide(
    "What AI Can Do",
    [
        "Summarise long documents into key points",
        "Translate text between languages (with imperfect nuance)",
        "Help plan projects, essays, or schedules",
        "Brainstorm ideas, names, titles, and arguments",
        "Generate templates: emails, outlines, cover letters",
        "Explain complex topics in simpler language",
        "Give feedback on writing drafts",
        "Answer general knowledge questions (with caveats)",
    ],
    subtitle="Tasks where AI adds genuine value",
    bullet_color=C_TEAL, icon="✓", accent=C_TEAL
)

# What AI CANNOT DO
bullet_slide(
    "What AI Cannot Do",
    [
        "Know the truth — AI can confidently state false information ('hallucinations')",
        "Be perfectly accurate — always verify important facts from primary sources",
        "Replace your judgment — AI has no understanding of your unique situation",
        "Understand cultural nuance, humor, sarcasm, or local context reliably",
        "Access live information (unless given web search tools explicitly)",
        "Remember previous conversations (by default)",
        "Think creatively — it recombines existing patterns, not novel ideas",
    ],
    subtitle="Critical limitations every user must know",
    bullet_color=C_PINK, icon="✗", accent=C_PINK
)

# Why AI Makes Mistakes
sl = add_slide()
header_bar(sl, "Why AI Makes Mistakes",
           subtitle="Understanding this makes you a smarter user",
           bar_color=C_ORANGE)
reasons = [
    ("💭 Hallucinations", C_PINK,
     "AI sometimes generates plausible-sounding but completely false information — "
     "including fake citations, statistics, or events."),
    ("⚖️ Bias in Training Data", C_ORANGE,
     "AI learns from internet text, which reflects human biases around gender, race, "
     "culture, and language. These biases can appear in outputs."),
    ("❓ Missing Context", C_BLUE,
     "Without enough context from you, AI fills gaps with assumptions — which may not "
     "match your actual situation or needs."),
    ("📅 Outdated Training Data", C_PURPLE,
     "Most AI models have a knowledge cutoff. Events, research, or policies after that "
     "date may not be reflected in responses."),
]
for i, (title, col, desc) in enumerate(reasons):
    row = i // 2; col_idx = i % 2
    bx = 0.3 + col_idx * 6.55; by = 1.58 + row * 2.85
    rect(sl, bx, by, 6.2, 2.65, col)
    txb(sl, title, bx + 0.2, by + 0.15, 5.8, 0.65,
        size=22, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.2), Inches(by + 0.85), Inches(5.8), Inches(1.65))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_WHITE

# Hallucination example
example_slide(
    "Hallucination Example",
    "Who wrote the novel 'The Language of Seasons' published in 2019?",
    "The novel 'The Language of Seasons' was written by Maria Chen and published by "
    "Penguin Random House in April 2019. It won the Booker Prize shortlist that year.\n\n"
    "⚠️  This book, author, and award are completely fabricated. "
    "Always verify book titles, authors, and awards with a library catalogue or publisher website.",
    subtitle="This is what a hallucination looks like — confident but false",
    accent=C_PINK
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2  –  MINDSET SHIFT
# ═════════════════════════════════════════════════════════════════════════════
section_divider(2, "Mindset Shift: Search → Collaboration",
                "Change how you talk to AI and transform what you get back",
                accent=C_TEAL)

bullet_slide(
    "The 'Search Engine' Mindset",
    [
        '"What is climate change?"',
        '"Causes of World War II"',
        '"Healthy foods"',
        '"Grammar rules English"',
        '"Essay about leadership"',
        '"Python tutorial"',
        '"Symptoms of anxiety"',
    ],
    subtitle="Short, keyword-style queries — passive information retrieval",
    bullet_color=C_PINK, icon="🔎", accent=C_PINK
)

bullet_slide(
    "The 'Collaborator' Mindset",
    [
        '"Act as an environmental scientist. Explain climate change to a first-year student in 3 bullet points."',
        '"Help me understand the main causes of WWI vs. WWII. Ask me questions first to check my knowledge gaps."',
        '"Give me a weekly meal plan as a table. I\'m vegetarian and on a budget."',
        '"I\'m writing an essay on leadership. Suggest 5 thesis angles and explain the pros of each."',
        '"I\'m a beginner. Teach me Python step-by-step, starting with variables. Check my understanding after each step."',
    ],
    subtitle="Specific, contextual, interactive — active collaboration",
    bullet_color=C_TEAL, icon="🤝", accent=C_TEAL
)

two_col_slide(
    "Search vs. Collaboration — Side by Side",
    "🔎  Search Mindset",
    [
        "Short, keyword queries",
        "Expects one correct answer",
        "Passive — reads and accepts",
        "No follow-up or iteration",
        "Treats AI like a dictionary",
        "Frustrated when answer is vague",
    ],
    "🤝  Collaborator Mindset",
    [
        "Context-rich, specific prompts",
        "Expects a starting point to refine",
        "Active — questions, pushes back",
        "Iterates and asks follow-ups",
        "Treats AI like a smart assistant",
        "Uses vague answers as a chance to clarify",
    ],
    left_col=C_PINK, right_col=C_TEAL,
    subtitle="Which mindset produces better results?"
)

activity_slide(
    "Mindset Makeover",
    [
        "Work with a partner (2 min each direction).",
        "",
        "Step 1:  Partner A reads a 'search mindset' prompt aloud.",
        "Step 2:  Partner B rewrites it as a 'collaborator mindset' prompt.",
        "Step 3:  Switch roles.",
        "",
        "Starter prompts to transform:",
        "→  'Study tips for exams'",
        "→  'How to write a conclusion'",
        "→  'Vocabulary for academic writing'",
        "→  'What is plagiarism?'",
        "",
        "Share your best rewrite with the group!",
    ],
    time_str="8 min", color=C_TEAL
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 3  –  PROMPTING LAB
# ═════════════════════════════════════════════════════════════════════════════
section_divider(3, "Prompting Lab",
                "4 Strategies  ·  Examples  ·  Try-It Activities",
                accent=C_PURPLE)

# ── Strategy 1: Role + Task ──────────────────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 1: Role + Task Prompting",
           subtitle="Tell AI WHO it is and WHAT you need it to do",
           bar_color=C_PURPLE)
rect(sl, 0.3, 1.58, 12.73, 5.6, C_LGRAY)
txb(sl, "TEMPLATE", 0.5, 1.68, 3.5, 0.45,
    size=15, bold=True, color=C_PURPLE)
txb(sl,
    '"Act as [ROLE]. Your audience is [AUDIENCE]. [TASK]. Use [TONE/FORMAT]."',
    0.5, 2.14, 12.3, 0.8,
    size=22, bold=True, color=C_PURPLE, italic=True)
rect(sl, 0.3, 3.05, 12.73, 0.06, C_PURPLE)
txb(sl, "Why it works:", 0.5, 3.15, 3.5, 0.42,
    size=17, bold=True, color=C_ORANGE)
reasons_s1 = [
    "▸  Role sets the expertise level and vocabulary AI uses",
    "▸  Task gives a clear goal — no guessing",
    "▸  Audience shapes complexity and tone",
    "▸  Format ensures output is ready to use",
]
by = 3.62
for r in reasons_s1:
    txb(sl, r, 0.65, by, 12.0, 0.46, size=19, color=C_DARK)
    by += 0.52

example_slide(
    "Strategy 1 — Role + Task Example",
    "Act as an academic English tutor. My student is a first-year international student "
    "who struggles with thesis statements. Explain what a thesis statement is and give 3 "
    "examples — one weak, one okay, one strong. Use simple language.",
    "A thesis statement is the main argument of your essay — it tells your reader what "
    "you will prove.\n\n"
    "Weak: 'Social media is popular.'  (too vague — everyone agrees)\n"
    "Okay: 'Social media affects students.'  (needs to say HOW and WHY)\n"
    "Strong: 'Excessive social media use reduces academic focus because it fragments "
    "attention and displaces study time.' ✓",
    subtitle="Role: Tutor  |  Audience: First-year student  |  Format: 3 examples",
    accent=C_PURPLE
)

activity_slide(
    "Try It: Role + Task",
    [
        "Write a Role + Task prompt for ONE of these situations:",
        "",
        "→  You are preparing for a job interview at a tech company.",
        "→  You need to explain your research topic to a non-expert friend.",
        "→  You want feedback on an email you wrote to a professor.",
        "→  You are writing a cover letter for an internship.",
        "",
        "Use the template:",
        "→  'Act as [ROLE]. My audience is [AUDIENCE]. [TASK]. Use [FORMAT].'",
        "",
        "Share your prompt with a neighbour and compare results!",
    ],
    time_str="6 min", color=C_PURPLE
)

# ── Strategy 2: Step-by-step / Scaffolded ────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 2: Step-by-Step / Scaffolded Prompting",
           subtitle="Break complex tasks into guided stages",
           bar_color=C_BLUE)
rect(sl, 0.3, 1.58, 12.73, 5.6, C_LGRAY)
txb(sl, "HOW TO ASK FOR BREAKDOWNS", 0.5, 1.68, 9, 0.45,
    size=16, bold=True, color=C_BLUE)
steps = [
    ("1", "Start Big",  "Ask for an overview or outline first."),
    ("2", "Zoom In",    "Follow up: 'Now explain step 2 in more detail.'"),
    ("3", "Check In",   "Ask: 'What have I missed? What should I consider next?'"),
    ("4", "Iterate",    "Refine: 'Simplify this section. Add an example.'"),
]
step_colors = [C_BLUE, C_PURPLE, C_TEAL, C_ORANGE]
for i, (num, title, desc) in enumerate(steps):
    bx = 0.4 + i * 3.22; by = 2.25
    rect(sl, bx, by, 3.05, 4.75, step_colors[i])
    txb(sl, num, bx + 0.12, by + 0.1, 0.78, 0.8,
        size=36, bold=True, color=C_WHITE)
    txb(sl, title, bx + 0.12, by + 1.02, 2.8, 0.58,
        size=20, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.12), Inches(by + 1.68), Inches(2.8), Inches(2.9))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_WHITE

example_slide(
    "Strategy 2 — Scaffolded Prompting Example",
    "I need to write a 1,500-word essay on globalisation and education. "
    "Start by giving me an outline with 5 sections and 2–3 bullet points per section. "
    "Do not write the essay yet — just the structure.",
    "Outline: Globalisation & Education\n"
    "1. Introduction  •  Define globalisation  •  Relevance to education  •  Thesis preview\n"
    "2. Benefits  •  Access to resources  •  Cultural exchange  •  International opportunities\n"
    "3. Challenges  •  Inequality of access  •  Language barriers  •  Cultural homogenisation\n"
    "4. Case Studies  •  Online learning platforms  •  International student mobility\n"
    "5. Conclusion  •  Summary  •  Future implications  •  Call to action",
    subtitle="Notice: you control the pace, AI builds step by step",
    accent=C_BLUE
)

activity_slide(
    "Try It: Scaffolded Prompting",
    [
        "Pick a task you are currently working on (or use the prompt below):",
        "",
        "→  'I need to study for my midterm exam on [your subject].'",
        "",
        "Write a scaffolded prompt sequence (3 prompts in a chain):",
        "→  Prompt 1: Ask for an overview / outline",
        "→  Prompt 2: Ask to expand on one section",
        "→  Prompt 3: Ask for a practice question or summary",
        "",
        "Optional: Try it live on your device and share the output!",
    ],
    time_str="7 min", color=C_BLUE
)

# ── Strategy 3: Format-based Output ──────────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 3: Format-Based Output",
           subtitle="Tell AI exactly how you want the answer structured",
           bar_color=C_TEAL)
formats = [
    ("📊 Tables",       "Compare options side by side.\n'Give me a table comparing X, Y, Z by cost, features, and ease of use.'"),
    ("☑️ Checklists",   "Get actionable to-do lists.\n'Give me a checklist of things to do before submitting my assignment.'"),
    ("📄 Templates",    "Generate reusable structures.\n'Create an email template I can use to ask professors for extensions.'"),
    ("• Bullet Points", "Scannable summaries.\n'Summarise this article in 5 bullet points, each under 15 words.'"),
    ("1. Numbered",     "Step-by-step instructions.\n'Give me numbered steps to set up a Python environment.'"),
    ("¶ Paragraphs",   "Flowing prose for essays.\n'Write this as 2 short paragraphs suitable for an academic essay.'"),
]
fmt_colors = [C_TEAL, C_BLUE, C_PURPLE, C_ORANGE, C_TEAL, C_BLUE]
for i, (fmt, desc) in enumerate(formats):
    row = i // 3; col_i = i % 3
    bx = 0.3 + col_i * 4.35; by = 1.62 + row * 2.82
    rect(sl, bx, by, 4.1, 2.62, fmt_colors[i])
    txb(sl, fmt, bx + 0.15, by + 0.14, 3.8, 0.58,
        size=19, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.15), Inches(by + 0.78), Inches(3.8), Inches(1.72))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(15); run.font.color.rgb = C_WHITE

example_slide(
    "Strategy 3 — Format-Based Output Example",
    "Compare three productivity apps (Notion, Todoist, Trello) using a table. "
    "Columns: App Name, Best For, Cost, Ease of Use (1–5 stars), Mobile App. "
    "Add a recommendation row at the bottom.",
    "| App      | Best For         | Cost        | Ease ⭐ | Mobile |\n"
    "|----------|-----------------|-------------|--------|--------|\n"
    "| Notion   | Notes + projects | Free / $8/mo |  ★★★★  |  ✓   |\n"
    "| Todoist  | Task management  | Free / $4/mo |  ★★★★★ |  ✓   |\n"
    "| Trello   | Visual boards    | Free / $5/mo |  ★★★★  |  ✓   |\n"
    "| 🏆 Pick  | Beginners → Todoist  |  Teams → Trello  |  Power users → Notion |",
    subtitle="Formatted output is immediately usable — no reformatting required",
    accent=C_TEAL
)

activity_slide(
    "Try It: Format-Based Output",
    [
        "Rewrite one of these generic prompts to include a format instruction:",
        "",
        "→  'Tips for managing stress during exams'",
        "→  'Differences between British and American English'",
        "→  'How to write a research paper'",
        "",
        "Choose a format:  Table  |  Checklist  |  Template  |  Numbered Steps",
        "",
        "Example rewrite:",
        "→  'Give me a checklist of 8 steps for managing exam stress, with a ☐ checkbox before each item.'",
        "",
        "Compare outputs with a partner who chose the same topic but a different format!",
    ],
    time_str="6 min", color=C_ORANGE
)

# ── Strategy 4: Revision & Reflection ────────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 4: Revision & Reflection Prompting",
           subtitle="Use AI as an editor, critic, and thinking partner",
           bar_color=C_ORANGE)
techniques = [
    ("✍️ Ask for Improvement", C_ORANGE,
     "Paste your draft and say:\n'Improve the clarity and flow of this paragraph. "
     "Keep my original voice.'"),
    ("❓ Ask for Questions First", C_PURPLE,
     "Before a big task say:\n'Before you help me, ask me 5 questions to better "
     "understand my needs and context.'"),
    ("🔍 Ask for Critique", C_PINK,
     "After a draft say:\n'What are the three weakest parts of this argument? "
     "How would you strengthen each one?'"),
    ("🔀 Ask for Alternatives", C_BLUE,
     "After an output say:\n'Give me 3 alternative ways to phrase this conclusion. "
     "Rate each one and explain why.'"),
]
for i, (title, col, desc) in enumerate(techniques):
    bx = 0.3 + (i % 2) * 6.55; by = 1.62 + (i // 2) * 2.82
    rect(sl, bx, by, 6.2, 2.62, col)
    txb(sl, title, bx + 0.2, by + 0.14, 5.8, 0.58,
        size=21, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.2), Inches(by + 0.8), Inches(5.8), Inches(1.68))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_WHITE

example_slide(
    "Strategy 4 — Revision Example",
    "Here is my conclusion paragraph: [paste paragraph]. "
    "Identify 2 weaknesses and suggest specific improvements. "
    "Then give me a revised version that is more academic in tone.",
    "Weaknesses identified:\n"
    "1. Your conclusion repeats the introduction almost word-for-word — "
    "it should synthesise rather than summarise.\n"
    "2. The final sentence is vague: 'AI is important for the future.' "
    "— add a specific implication.\n\n"
    "Revised conclusion:\n"
    "'Taken together, these findings suggest that AI literacy is no longer optional "
    "for university students — it is a foundational skill for participation in the "
    "modern knowledge economy.'",
    subtitle="AI as editor — keeps your voice, strengthens your argument",
    accent=C_ORANGE
)

activity_slide(
    "Try It: Revision & Reflection",
    [
        "Choose one of the following activities:",
        "",
        "Option A — Ask questions first:",
        "→  Type: 'I want help planning my next assignment. Ask me 5 questions first.'",
        "→  Answer the questions, then see how the AI's help improves.",
        "",
        "Option B — Revision loop:",
        "→  Paste a sentence or paragraph you wrote recently.",
        "→  Ask AI to identify weaknesses and offer a revision.",
        "→  Then ask: 'Now make it 20% shorter without losing meaning.'",
        "",
        "Reflect: How did the output change when you gave more context?",
    ],
    time_str="8 min", color=C_PURPLE
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 4  –  EVALUATING AI OUTPUT
# ═════════════════════════════════════════════════════════════════════════════
section_divider(4, "Evaluating AI Output",
                "Not all AI answers are equal — learn to tell the difference",
                accent=C_ORANGE)

# Why different tools give different answers
sl = add_slide()
header_bar(sl, "Why Different AI Tools Give Different Answers",
           subtitle="Same question, different outputs — here's why",
           bar_color=C_ORANGE)
reasons_t = [
    ("🗂️ Different Training Data",
     "Each AI model is trained on different datasets with different time cutoffs, "
     "languages, and sources of information."),
    ("🏗️ Different Model Architecture",
     "GPT-4, Gemini, Claude, and Llama use different underlying approaches to "
     "language modelling — producing different styles and strengths."),
    ("🔧 Different System Prompts",
     "Every AI product has hidden instructions shaping tone, safety limits, "
     "and focus areas — you don't always see these."),
    ("🎲 Randomness (Temperature)",
     "AI responses include controlled randomness. Run the same prompt twice "
     "and you may get different outputs."),
    ("📏 Context Length",
     "Some tools handle longer conversations or documents better than others, "
     "affecting quality in extended tasks."),
]
by = 1.65
bar_colors = [C_ORANGE, C_PURPLE, C_BLUE, C_TEAL, C_PINK]
for idx, (r_title, r_desc) in enumerate(reasons_t):
    rect(sl, 0.4, by, 12.53, 1.04, bar_colors[idx])
    txb(sl, r_title, 0.62, by + 0.1, 4.8, 0.48,
        size=18, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(Inches(5.6), Inches(by + 0.1), Inches(7.2), Inches(0.84))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = r_desc; run.font.size = Pt(16); run.font.color.rgb = C_WHITE
    by += 1.1

# Evaluation checklist
sl = add_slide()
header_bar(sl, "AI Output Evaluation Checklist",
           subtitle="Use this every time before you trust or submit AI content",
           bar_color=C_BLUE)
checklist = [
    ("☐  Accuracy",           C_PINK,    "Can you verify key facts from a reliable primary source?"),
    ("☐  Bias",               C_ORANGE,  "Does the response favour one perspective unfairly?"),
    ("☐  Missing Citations",  C_PURPLE,  "Are sources named? Are they real and accessible?"),
    ("☐  Clarity",            C_BLUE,    "Is the language clear, relevant, and appropriate for your audience?"),
    ("☐  Cultural Relevance", C_TEAL,    "Does it account for your cultural or regional context?"),
    ("☐  Completeness",       C_GREEN,   "Does it fully address your question, or are there gaps?"),
]
for i, (item, col, desc) in enumerate(checklist):
    by = 1.58 + i * 0.98
    rect(sl, 0.3, by, 12.73, 0.9, col)
    txb(sl, item, 0.5, by + 0.14, 4.4, 0.62,
        size=20, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(Inches(5.1), Inches(by + 0.14), Inches(7.8), Inches(0.65))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_WHITE

# Practice — spot the errors
sl = add_slide()
header_bar(sl, "Practice: Spot the Errors",
           subtitle="Read this AI output — how many problems can you find?",
           bar_color=C_PINK)
rect(sl, 0.3, 1.58, 12.73, 4.15, C_LGRAY)
txb(sl, "AI OUTPUT (contains deliberate errors):", 0.5, 1.68, 10, 0.45,
    size=15, bold=True, color=C_PINK)
error_text = (
    "The United Nations was founded in 1945 by Franklin D. Roosevelt and Winston Churchill "
    "at the Geneva Conference. It currently has 184 member states. The UN's primary goal "
    "is to prevent wars and promote peace worldwide. Its headquarters are located in Paris, "
    "France. The Secretary-General serves a 3-year term. The UN Security Council has "
    "7 permanent members including the USA, UK, France, Russia, and China."
)
box = sl.shapes.add_textbox(Inches(0.5), Inches(2.18), Inches(12.3), Inches(3.35))
tf = box.text_frame; tf.word_wrap = True
run = tf.paragraphs[0].add_run()
run.text = error_text; run.font.size = Pt(20); run.font.color.rgb = C_DARK; run.font.italic = True

rect(sl, 0.3, 5.85, 12.73, 1.42, C_ORANGE)
txb(sl, "❓  Errors to find — discuss with a partner, then reveal!",
    0.5, 5.92, 12.0, 0.42, size=16, bold=True, color=C_WHITE)
answers = ("Answers: ① Geneva → San Francisco  "
           "② 184 → 193 members  "
           "③ Paris → New York  "
           "④ 3-year → 5-year term  "
           "⑤ 7 → 5 permanent members")
txb(sl, answers, 0.5, 6.4, 12.5, 0.8, size=15, color=C_WHITE, bold=True)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 5  –  AI GUIDELINES / RESPONSIBLE USE
# ═════════════════════════════════════════════════════════════════════════════
section_divider(5, "AI Guidelines & Responsible Use",
                "Know the rules — protect yourself and others",
                accent=C_GREEN)

bullet_slide(
    "When AI IS Allowed in Academic Settings",
    [
        "Brainstorming and generating initial ideas (with disclosure)",
        "Improving grammar and clarity of your own draft",
        "Translating content to better understand a source",
        "Summarising readings to aid comprehension",
        "Generating practice questions or study guides",
        "Getting feedback on structure or argument strength",
        "Learning new concepts through interactive explanation",
    ],
    subtitle="Always check your institution's specific policy first",
    bullet_color=C_TEAL, icon="✓", accent=C_TEAL
)

bullet_slide(
    "When AI is NOT Allowed",
    [
        "Generating entire assignments or essays for submission",
        "Completing exams, quizzes, or timed assessments",
        "Writing lab reports, reflections, or personal statements",
        "Paraphrasing AI content to disguise AI origin",
        "Using AI to complete peer evaluations or feedback forms",
        "Any situation where your instructor has explicitly prohibited it",
        "Presenting AI output as entirely your own original thinking",
    ],
    subtitle="Academic integrity applies to AI — misuse has serious consequences",
    bullet_color=C_PINK, icon="✗", accent=C_PINK
)

# Privacy + safety
sl = add_slide()
header_bar(sl, "Privacy, Safety & Personal Data",
           subtitle="What you share with AI is not always private",
           bar_color=C_PURPLE)
privacy_items = [
    ("🔒 Never Share", C_PINK, [
        "Passport, ID, or social security numbers",
        "Bank details, card numbers, or financial data",
        "Medical records or health information",
        "Passwords or account credentials",
        "Classmates' personal information or work",
    ]),
    ("⚠️ Be Cautious With", C_ORANGE, [
        "Your full name combined with location",
        "Details about your mental health or personal struggles",
        "Confidential course materials (copyrighted content)",
        "Sensitive family or relationship information",
        "Unpublished research data",
    ]),
]
for i, (heading, col, items) in enumerate(privacy_items):
    bx = 0.4 + i * 6.55
    rect(sl, bx, 1.62, 6.2, 5.58, col)
    txb(sl, heading, bx + 0.2, 1.74, 5.8, 0.68,
        size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    box = sl.shapes.add_textbox(Inches(bx + 0.2), Inches(2.52), Inches(5.8), Inches(4.4))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for item in items:
        if first: p = tf.paragraphs[0]; first = False
        else: p = tf.add_paragraph()
        p.space_before = Pt(9)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(18); run.font.color.rgb = C_WHITE

two_col_slide(
    "Ethical Use: Learning vs. Shortcuts",
    "✅  Using AI to LEARN",
    [
        "Ask AI to explain a concept, then test yourself",
        "Use AI feedback to improve your own draft",
        "Ask AI questions to deepen understanding",
        "Use AI to explore multiple perspectives",
        "Cite AI assistance when required",
        "Reflect on what AI taught you",
    ],
    "❌  Using AI as a SHORTCUT",
    [
        "Submitting AI output without reading it",
        "Replacing your thinking with AI thinking",
        "Hiding your AI use from instructors",
        "Copying AI text and calling it your writing",
        "Using AI to avoid engaging with material",
        "Letting AI do the learning for you",
    ],
    left_col=C_TEAL, right_col=C_PINK,
    subtitle="The difference is whether YOU are growing"
)

bullet_slide(
    "Basic AI Citation Expectations",
    [
        "Check your institution's style guide — many now include AI citation formats.",
        "APA 7:  OpenAI. (2024). ChatGPT (GPT-4) [Large language model]. https://chat.openai.com",
        "MLA:  'Prompt text.' ChatGPT, GPT-4 version, OpenAI, 15 Jan. 2024, chat.openai.com.",
        "Include: AI tool name  ·  Version  ·  Date of conversation  ·  Prompt used (if required)",
        "When in doubt, be transparent — disclose AI use in a brief author's note.",
        "Remember: citations show intellectual honesty, not weakness.",
    ],
    subtitle="When disclosure is required, cite clearly and completely",
    bullet_color=C_DGRAY, icon="📎", accent=C_BLUE
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 6  –  FINAL CHALLENGE + REFLECTION
# ═════════════════════════════════════════════════════════════════════════════
section_divider(6, "Final Prompting Challenge + Reflection",
                "Put all 4 strategies together in one powerful prompt",
                accent=C_ORANGE)

# Build-your-best-prompt
sl = add_slide()
header_bar(sl, "Build Your Best Prompt Challenge",
           subtitle="Combine all 6 elements into one masterful prompt",
           bar_color=C_ORANGE)
elements = [
    ("1", "Role",          "Who should AI be?",              "Act as a …"),
    ("2", "Context",       "What's the background?",         "I am a … working on …"),
    ("3", "Task",          "What exactly do you need?",      "Help me … by …"),
    ("4", "Constraints",   "What limits apply?",             "Keep it under … / Do not …"),
    ("5", "Output Format", "How should the answer look?",    "Format as a table / checklist …"),
    ("6", "Language Level","What reading level / language?", "Use B2 English / avoid jargon …"),
]
elem_colors = [C_PURPLE, C_BLUE, C_TEAL, C_ORANGE, C_PINK, C_GREEN]
for i, (num, label, question, hint) in enumerate(elements):
    row = i // 3; col_i = i % 3
    bx = 0.25 + col_i * 4.36; by = 1.58 + row * 2.85
    rect(sl, bx, by, 4.1, 2.68, elem_colors[i])
    txb(sl, num, bx + 0.12, by + 0.12, 0.7, 0.65,
        size=28, bold=True, color=C_WHITE)
    txb(sl, label, bx + 0.85, by + 0.12, 3.1, 0.62,
        size=20, bold=True, color=C_WHITE)
    txb(sl, question, bx + 0.12, by + 0.88, 3.88, 0.55,
        size=14, color=C_WHITE, italic=True)
    txb(sl, hint, bx + 0.12, by + 1.5, 3.88, 0.9,
        size=14, color=C_YELLOW, bold=True)

activity_slide(
    "Build Your Best Prompt — Live Challenge",
    [
        "Choose a real task you need help with this week.",
        "",
        "Write ONE prompt that includes ALL 6 elements:",
        "→  Role  |  Context  |  Task  |  Constraints  |  Output Format  |  Language Level",
        "",
        "Example master prompt:",
        "→  'Act as a career counsellor [ROLE]. I am an international student in my final year "
        "studying Computer Science [CONTEXT]. Help me write a cover letter for a software "
        "internship at Google [TASK]. Keep it under 300 words and avoid clichés [CONSTRAINTS]. "
        "Format it as a professional letter [FORMAT]. Use clear, confident B2 English [LANGUAGE].'",
        "",
        "Share your prompt — class votes on the most complete one! 🏆",
    ],
    time_str="10 min", color=C_ORANGE
)

# Reflection
sl = add_slide()
header_bar(sl, "Reflection: Mindset Shift & Next Steps",
           subtitle="Take 3 minutes to think — then share one insight",
           bar_color=C_PURPLE)
rect(sl, 0.3, 1.58, 12.73, 5.62, C_LGRAY)
reflections = [
    ("🔄", "Mindset Shift",   "At the start of today, how did you think about AI? How has that changed?"),
    ("💡", "Biggest Insight", "What is the single most useful thing you learned in this workshop?"),
    ("✏️", "First Prompt",    "What is the first real prompt you will try using today's strategies?"),
    ("⚠️", "Responsible Use", "What is one way you will use AI MORE responsibly going forward?"),
    ("🚀", "Next Step",       "What skill do you want to develop further — prompting, evaluation, or something else?"),
]
refl_colors = [C_PURPLE, C_BLUE, C_TEAL, C_ORANGE, C_PINK]
by = 1.68
for idx, (emoji, title, question) in enumerate(reflections):
    rect(sl, 0.4, by, 12.53, 1.02, refl_colors[idx])
    txb(sl, f"{emoji}  {title}", 0.62, by + 0.1, 3.8, 0.52,
        size=18, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(Inches(4.6), Inches(by + 0.12), Inches(8.1), Inches(0.78))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = question; run.font.size = Pt(17); run.font.color.rgb = C_WHITE
    by += 1.08

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 7  –  CLOSING
# ═════════════════════════════════════════════════════════════════════════════
section_divider(7, "Closing & Resources",
                "Thank you — now go collaborate!",
                accent=C_TEAL)

sl = add_slide(bg=C_PURPLE)
rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
rect(sl, 0.55, 0, 13.33, 7.5, C_PURPLE)
txb(sl, "Thank You!", 1.0, 0.5, 11.5, 1.05,
    size=52, bold=True, color=C_YELLOW)
txb(sl, "You are now an AI collaborator — not just an AI user.",
    1.0, 1.6, 11.5, 0.72, size=26, color=C_WHITE, italic=True)

rect(sl, 0.9, 2.5, 11.53, 4.65, RGBColor(0x5A, 0x28, 0xC8))
txb(sl, "Resources & Prompt Templates", 1.1, 2.6, 11.0, 0.65,
    size=22, bold=True, color=C_YELLOW)
resources = [
    "🔗  ChatGPT: chat.openai.com         🔗  Claude: claude.ai",
    "🔗  Gemini:  gemini.google.com        🔗  Copilot: copilot.microsoft.com",
    "",
    "📋  Master Prompt Template:",
    "     'Act as [ROLE]. I am [CONTEXT]. Help me [TASK]. Keep [CONSTRAINTS]. Format as [FORMAT]. Use [LANGUAGE].'",
    "",
    "📚  Further Learning:  learnprompting.org  ·  elementsofai.com  ·  promptingguide.ai",
    "",
    "📧  Workshop Support:  [facilitator@institution.edu]",
    "📞  Student Services:  [Your institution's student support contact]",
]
box = sl.shapes.add_textbox(Inches(1.1), Inches(3.3), Inches(11.1), Inches(3.65))
tf = box.text_frame; tf.word_wrap = True
first = True
for line in resources:
    if first: p = tf.paragraphs[0]; first = False
    else: p = tf.add_paragraph()
    p.space_before = Pt(5)
    run = p.add_run(); run.text = line
    run.font.size = Pt(16)
    if line.startswith("📋") or line.startswith("📚"):
        run.font.color.rgb = C_YELLOW; run.font.bold = True
    elif line.startswith("     '"):
        run.font.color.rgb = C_LGRAY; run.font.italic = True
    else:
        run.font.color.rgb = C_LGRAY

txb(sl, "Remember: The best prompt is the one that helps you think more clearly.",
    0.9, 7.08, 11.53, 0.4, size=15, color=C_YELLOW, italic=True,
    align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/powerpoints/AI_Workshop_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}  ({len(prs.slides)} slides)")
