from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Palette ──────────────────────────────────────────────────────────────────
C_DARK   = RGBColor(0x1A, 0x1A, 0x2E)   # deep navy  – slide backgrounds
C_MID    = RGBColor(0x16, 0x21, 0x3E)   # mid navy   – accent panels
C_ACCENT = RGBColor(0x0F, 0x3D, 0x96)   # royal blue – headers / borders
C_TEAL   = RGBColor(0x00, 0xB4, 0xD8)   # cyan‑teal  – highlights
C_GREEN  = RGBColor(0x06, 0xD6, 0xA0)   # mint green – positive / correct
C_ORANGE = RGBColor(0xFF, 0x93, 0x00)   # amber      – warnings / activities
C_RED    = RGBColor(0xEF, 0x47, 0x6F)   # rose‑red   – incorrect / caution
C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
C_LGRAY  = RGBColor(0xD0, 0xD8, 0xE8)   # light grey – body text
C_YELLOW = RGBColor(0xFF, 0xD1, 0x66)   # soft gold  – quiz / emphasis

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]   # completely blank

# ─────────────────────────────────────────────────────────────────────────────
# Helper utilities
# ─────────────────────────────────────────────────────────────────────────────

def add_slide():
    sl = prs.slides.add_slide(BLANK)
    # dark background
    fill = sl.background.fill
    fill.solid()
    fill.fore_color.rgb = C_DARK
    return sl


def txb(sl, text, x, y, w, h,
        size=24, bold=False, color=C_WHITE, align=PP_ALIGN.LEFT,
        italic=False, wrap=True):
    """Add a text‑box and return the run so caller can tweak further."""
    box  = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf   = box.text_frame
    tf.word_wrap = wrap
    p    = tf.paragraphs[0]
    p.alignment = align
    run  = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic    = italic
    return run


def para(tf, text, size=20, bold=False, color=C_LGRAY,
         align=PP_ALIGN.LEFT, italic=False, space_before=6):
    """Append a paragraph to an existing text‑frame."""
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = Pt(space_before)
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.bold  = bold
    run.font.color.rgb = color
    run.font.italic    = italic
    return p


def rect(sl, x, y, w, h, fill_color, alpha=None):
    """Solid filled rectangle (no outline)."""
    shape = sl.shapes.add_shape(
        1,   # MSO_SHAPE_TYPE.RECTANGLE
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def header_bar(sl, title, subtitle=None,
               bar_color=C_ACCENT, title_color=C_WHITE, sub_color=C_TEAL):
    """Top banner with title (and optional subtitle)."""
    rect(sl, 0, 0, 13.33, 1.35, bar_color)
    txb(sl, title, 0.3, 0.1, 12.5, 0.75,
        size=36, bold=True, color=title_color, align=PP_ALIGN.LEFT)
    if subtitle:
        txb(sl, subtitle, 0.35, 0.85, 12.5, 0.5,
            size=18, color=sub_color, italic=True)


def section_divider(section_num, section_title, subtitle=""):
    """Full‑bleed section title slide."""
    sl = add_slide()
    # accent stripe on left
    rect(sl, 0, 0, 0.4, 7.5, C_TEAL)
    # section number badge
    badge = sl.shapes.add_shape(1, Inches(1.5), Inches(2.5), Inches(1.8), Inches(1.8))
    badge.fill.solid(); badge.fill.fore_color.rgb = C_ACCENT
    badge.line.fill.background()
    txb(sl, str(section_num), 1.5, 2.5, 1.8, 1.8,
        size=64, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)
    txb(sl, section_title, 3.6, 2.55, 9.0, 1.1,
        size=42, bold=True, color=C_WHITE)
    if subtitle:
        txb(sl, subtitle, 3.6, 3.75, 9.0, 0.9,
            size=22, color=C_LGRAY, italic=True)
    return sl


def bullet_slide(title, bullets, subtitle=None,
                 bullet_color=C_LGRAY, icon="▸", size=21):
    """Standard bullet‑list slide."""
    sl = add_slide()
    header_bar(sl, title, subtitle)
    box = sl.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12.3), Inches(5.6))
    tf  = box.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(8)
        run = p.add_run()
        run.text = f"{icon}  {b}" if icon else b
        run.font.size  = Pt(size)
        run.font.color.rgb = bullet_color
    return sl


def two_col_slide(title, left_title, left_items,
                  right_title, right_items,
                  left_col=C_ACCENT, right_col=C_MID,
                  subtitle=None):
    """Two‑column comparison slide."""
    sl = add_slide()
    header_bar(sl, title, subtitle)
    # left panel
    rect(sl, 0.3, 1.5, 6.1, 5.6, left_col)
    txb(sl, left_title, 0.5, 1.6, 5.7, 0.7,
        size=22, bold=True, color=C_YELLOW, align=PP_ALIGN.CENTER)
    box_l = sl.shapes.add_textbox(Inches(0.5), Inches(2.4), Inches(5.7), Inches(4.5))
    tf_l  = box_l.text_frame; tf_l.word_wrap = True
    first = True
    for item in left_items:
        if first: p = tf_l.paragraphs[0]; first = False
        else: p = tf_l.add_paragraph()
        p.space_before = Pt(7)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(19); run.font.color.rgb = C_WHITE
    # right panel
    rect(sl, 6.9, 1.5, 6.1, 5.6, right_col)
    txb(sl, right_title, 7.1, 1.6, 5.7, 0.7,
        size=22, bold=True, color=C_GREEN, align=PP_ALIGN.CENTER)
    box_r = sl.shapes.add_textbox(Inches(7.1), Inches(2.4), Inches(5.7), Inches(4.5))
    tf_r  = box_r.text_frame; tf_r.word_wrap = True
    first = True
    for item in right_items:
        if first: p = tf_r.paragraphs[0]; first = False
        else: p = tf_r.add_paragraph()
        p.space_before = Pt(7)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(19); run.font.color.rgb = C_WHITE
    return sl


def activity_slide(title, instructions, time_str="5 min", color=C_ORANGE):
    """Hands‑on activity slide."""
    sl = add_slide()
    # banner
    rect(sl, 0, 0, 13.33, 1.35, color)
    txb(sl, f"✏️  Activity: {title}", 0.3, 0.1, 11.5, 0.75,
        size=32, bold=True, color=C_DARK)
    txb(sl, f"⏱  {time_str}", 11.0, 0.1, 2.0, 0.7,
        size=22, bold=True, color=C_DARK, align=PP_ALIGN.RIGHT)
    # instructions box
    rect(sl, 0.5, 1.6, 12.3, 5.5, C_MID)
    box = sl.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.8), Inches(5.1))
    tf  = box.text_frame; tf.word_wrap = True
    first = True
    for line in instructions:
        if first: p = tf.paragraphs[0]; first = False
        else: p = tf.add_paragraph()
        p.space_before = Pt(9)
        run = p.add_run(); run.text = line
        run.font.size = Pt(21)
        run.font.color.rgb = C_WHITE if not line.startswith("→") else C_TEAL
    return sl


def example_slide(title, prompt_text, output_text, subtitle=None):
    """Example prompt / AI‑output slide."""
    sl = add_slide()
    header_bar(sl, title, subtitle)
    # prompt box
    rect(sl, 0.3, 1.55, 12.7, 0.4, C_TEAL)
    txb(sl, "PROMPT", 0.5, 1.57, 3, 0.35,
        size=14, bold=True, color=C_DARK)
    rect(sl, 0.3, 1.95, 12.7, 2.0, C_MID)
    box_p = sl.shapes.add_textbox(Inches(0.5), Inches(2.0), Inches(12.2), Inches(1.85))
    tf_p  = box_p.text_frame; tf_p.word_wrap = True
    run = tf_p.paragraphs[0].add_run()
    run.text = prompt_text
    run.font.size = Pt(19); run.font.color.rgb = C_YELLOW; run.font.italic = True
    # output box
    rect(sl, 0.3, 4.1, 12.7, 0.4, C_GREEN)
    txb(sl, "AI OUTPUT", 0.5, 4.12, 4, 0.35,
        size=14, bold=True, color=C_DARK)
    rect(sl, 0.3, 4.5, 12.7, 2.65, C_MID)
    box_o = sl.shapes.add_textbox(Inches(0.5), Inches(4.55), Inches(12.2), Inches(2.5))
    tf_o  = box_o.text_frame; tf_o.word_wrap = True
    run2 = tf_o.paragraphs[0].add_run()
    run2.text = output_text
    run2.font.size = Pt(18); run2.font.color.rgb = C_LGRAY
    return sl


def quiz_question_slide(q_num, question, options, note=None):
    """Multiple‑choice quiz question slide."""
    sl = add_slide()
    rect(sl, 0, 0, 13.33, 1.0, C_YELLOW)
    txb(sl, f"Q{q_num}", 0.3, 0.12, 1.2, 0.75,
        size=36, bold=True, color=C_DARK)
    txb(sl, "Opening Quiz", 1.6, 0.22, 10, 0.6,
        size=22, bold=True, color=C_DARK)
    txb(sl, question, 0.4, 1.1, 12.5, 1.1,
        size=26, bold=True, color=C_WHITE)
    letters = ["A", "B", "C", "D"]
    colors  = [C_ACCENT, C_MID, C_ACCENT, C_MID]
    for i, opt in enumerate(options):
        row = i // 2; col = i % 2
        bx = 0.4 + col * 6.5; by = 2.5 + row * 1.9
        rect(sl, bx, by, 6.2, 1.6, colors[i])
        txb(sl, letters[i], bx + 0.15, by + 0.2, 0.7, 1.1,
            size=30, bold=True, color=C_TEAL)
        box = sl.shapes.add_textbox(
            Inches(bx + 0.9), Inches(by + 0.2), Inches(5.1), Inches(1.2))
        tf = box.text_frame; tf.word_wrap = True
        run = tf.paragraphs[0].add_run()
        run.text = opt
        run.font.size = Pt(19); run.font.color.rgb = C_WHITE
    if note:
        txb(sl, f"💬  {note}", 0.4, 7.0, 12.5, 0.45,
            size=14, color=C_LGRAY, italic=True)
    return sl


# ═════════════════════════════════════════════════════════════════════════════
# SLIDE 1  –  Title / Cover
# ═════════════════════════════════════════════════════════════════════════════
sl = add_slide()
rect(sl, 0, 0, 13.33, 7.5, C_DARK)
# decorative accent lines
rect(sl, 0, 0, 0.5, 7.5, C_TEAL)
rect(sl, 0.5, 0, 0.15, 7.5, C_ACCENT)
# title
txb(sl, "AI as Your Collaborator", 1.0, 1.4, 11.8, 1.5,
    size=52, bold=True, color=C_WHITE)
txb(sl, "An Interactive Workshop for International College Students",
    1.0, 3.05, 11.5, 0.8, size=26, color=C_TEAL, italic=True)
# meta info row
rect(sl, 1.0, 4.1, 11.0, 0.08, C_ACCENT)
txb(sl, "Basics  ·  Mindset Shift  ·  Prompting Lab  ·  Evaluation  ·  Responsible Use",
    1.0, 4.3, 11.5, 0.6, size=18, color=C_LGRAY)
txb(sl, "Facilitator Guide  |  Workshop Version 1.0",
    1.0, 5.2, 8, 0.5, size=16, color=C_LGRAY, italic=True)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 0  –  OPENING QUIZ
# ═════════════════════════════════════════════════════════════════════════════
section_divider(0, "Opening Quiz",
                "Let's see what you already know — no grades, just fun!")

# --- Quiz Instructions ---
sl = add_slide()
header_bar(sl, "How the Quiz Works", subtitle="Read each question — answer in your head or on paper",
           bar_color=C_YELLOW, title_color=C_DARK, sub_color=C_DARK)
rect(sl, 0.5, 1.55, 12.3, 5.6, C_MID)
box = sl.shapes.add_textbox(Inches(0.8), Inches(1.75), Inches(11.8), Inches(5.2))
tf  = box.text_frame; tf.word_wrap = True
instructions = [
    "📋  There are 12 quick questions across 4 categories.",
    "🤔  Answer honestly — this is NOT graded.",
    "🤫  Keep your answers private for now; we'll discuss after.",
    "📱  You may use any device to jot notes, or just think silently.",
    "⏱   We will spend about 1–2 minutes per question.",
    "🎯  Goal: understand where you are starting from today.",
    "",
    "Categories:   AI Basics  |  Mindset  |  Prompting Clarity  |  Responsible Use",
]
first = True
for line in instructions:
    if first: p = tf.paragraphs[0]; first = False
    else: p = tf.add_paragraph()
    p.space_before = Pt(6)
    run = p.add_run(); run.text = line
    run.font.size = Pt(20)
    run.font.color.rgb = C_YELLOW if line.startswith("Categories") else C_WHITE

# --- AI Basics questions ---
quiz_question_slide(1,
    "What does AI stand for?",
    ["Automatic Internet", "Artificial Intelligence",
     "Advanced Information", "Automated Input"],
    note="Think about what each word means individually.")

quiz_question_slide(2,
    "AI chatbots like ChatGPT generate responses by…",
    ["Searching the live internet for every answer",
     "Predicting the next likely word based on patterns",
     "Consulting a database of correct answers",
     "Having human experts review each response"],
    note="Consider how 'language models' actually work.")

quiz_question_slide(3,
    "Which of the following is something AI does BEST?",
    ["Verifying that facts are 100% accurate",
     "Understanding cultural nuance and humor",
     "Generating text drafts and summarizing content",
     "Replacing human creative judgment"],
    note="Think about real‑world AI use cases.")

# --- Mindset shift questions ---
quiz_question_slide(4,
    "Which prompt is more likely to give a useful AI response?",
    ["What is climate change?",
     "Help me write a 3‑sentence explanation of climate change for a 10‑year‑old.",
     "Climate change",
     "Tell me about climate."],
    note="Think about what gives the AI the most context.")

quiz_question_slide(5,
    "You ask AI a question and the answer seems wrong. What do you do?",
    ["Trust it — AI is always accurate",
     "Give up and try Google",
     "Follow up with a new prompt asking it to reconsider",
     "Copy and submit the answer as‑is"],
    note="This is about thinking like a collaborator.")

# --- Prompting Clarity questions ---
quiz_question_slide(6,
    "What does adding a ROLE to a prompt do?",
    ["Makes the AI respond faster",
     "Tells the AI what persona or expertise to use",
     "Prevents the AI from making mistakes",
     "Limits how long the response will be"],
    note="Example: 'Act as an academic writing tutor…'")

quiz_question_slide(7,
    "Which prompt would produce the most structured output?",
    ["Tell me about healthy eating.",
     "What should I eat?",
     "Give me a weekly meal plan as a table with breakfast, lunch, and dinner.",
     "List foods."],
    note="Think about how format instructions shape AI output.")

quiz_question_slide(8,
    "Why might you ask AI to 'ask you questions first'?",
    ["To slow the AI down",
     "To avoid getting any response",
     "So the AI can gather context before giving advice",
     "Because AI cannot answer without questions"],
    note="This is a powerful prompting technique!")

# --- Appropriate vs. inappropriate AI use ---
quiz_question_slide(9,
    "Which is an APPROPRIATE use of AI in academic work?",
    ["Having AI write your entire essay and submitting it",
     "Using AI to brainstorm ideas, then writing in your own words",
     "Copying AI output without citing it",
     "Asking AI for your exam answers"],
    note="Know your institution's AI policy.")

quiz_question_slide(10,
    "What should you NEVER share with an AI chatbot?",
    ["Your essay topic",
     "A word you want to define",
     "Your passport number or bank details",
     "Your hometown"],
    note="Privacy matters — AI conversations may be stored.")

quiz_question_slide(11,
    "AI makes mistakes because…",
    ["It is lazy",
     "Developers program it to be wrong",
     "Its training data may be incomplete, biased, or outdated",
     "It only works in English"],
    note="Understanding AI limits makes you a smarter user.")

quiz_question_slide(12,
    "Which best describes using AI ETHICALLY as a student?",
    ["Using AI to learn, draft, and improve your thinking",
     "Replacing all your own work with AI output",
     "Sharing your classmates' work with AI",
     "Using AI to cheat in exams"],
    note="Ethics = using AI to grow, not to avoid growing.")

# --- Quiz Debrief ---
sl = add_slide()
header_bar(sl, "Quiz Debrief", subtitle="Let's talk about what we discovered",
           bar_color=C_GREEN, title_color=C_DARK, sub_color=C_DARK)
rect(sl, 0.5, 1.55, 12.3, 5.6, C_MID)
box = sl.shapes.add_textbox(Inches(0.8), Inches(1.75), Inches(11.8), Inches(5.2))
tf  = box.text_frame; tf.word_wrap = True
debrief = [
    "✅  Q1–Q3  |  AI Basics:  AI predicts language; it doesn't 'know' truth.",
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
    p.space_before = Pt(8)
    run = p.add_run(); run.text = line
    run.font.size = Pt(20)
    if line.startswith("→"):
        run.font.color.rgb = C_TEAL; run.font.bold = True
    elif line.startswith("Key"):
        run.font.color.rgb = C_YELLOW; run.font.bold = True
    else:
        run.font.color.rgb = C_WHITE

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 1  –  BASICS OF AI
# ═════════════════════════════════════════════════════════════════════════════
section_divider(1, "Basics of AI",
                "What it is, what it can do, and where it falls short")

# What AI IS
sl = add_slide()
header_bar(sl, "What Is AI?", subtitle="A simple explanation for everyday use")
# 3 concept boxes
concepts = [
    ("Pattern Recognition", C_ACCENT,
     "AI learns by finding patterns in enormous amounts of text, images, and data. "
     "It does not 'think' — it matches patterns."),
    ("Language Generation", C_MID,
     "AI chatbots predict the most likely next word — like autocomplete on a massive scale. "
     "Each word is chosen statistically."),
    ("No Understanding", RGBColor(0x2D, 0x00, 0x5E),
     "AI does not truly 'know' things. It has no beliefs, no memory between sessions "
     "(by default), and no lived experience."),
]
for i, (title, col, desc) in enumerate(concepts):
    bx = 0.3 + i * 4.35; by = 1.55
    rect(sl, bx, by, 4.1, 5.6, col)
    txb(sl, title, bx + 0.15, by + 0.2, 3.8, 0.75,
        size=21, bold=True, color=C_TEAL, align=PP_ALIGN.CENTER)
    rect(sl, bx + 0.1, by + 1.0, 3.9, 0.06, C_TEAL)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.15), Inches(by + 1.15), Inches(3.8), Inches(4.1))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(18); run.font.color.rgb = C_LGRAY

# analogy footer
txb(sl, "🔍  Analogy: AI is like a very well‑read autocomplete — not a search engine, not a brain.",
    0.3, 6.9, 12.7, 0.5, size=16, color=C_TEAL, italic=True)

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
    bullet_color=C_GREEN,
    icon="✓"
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
    bullet_color=C_RED,
    icon="✗"
)

# Why AI Makes Mistakes
sl = add_slide()
header_bar(sl, "Why AI Makes Mistakes",
           subtitle="Understanding this makes you a smarter user")
reasons = [
    ("Hallucinations", C_RED,
     "AI sometimes generates plausible‑sounding but completely false information — "
     "including fake citations, statistics, or events."),
    ("Bias in Training Data", C_ORANGE,
     "AI learns from internet text, which reflects human biases around gender, race, "
     "culture, and language. These biases can appear in outputs."),
    ("Missing Context", C_ACCENT,
     "Without enough context from you, AI fills gaps with assumptions — which may not "
     "match your actual situation or needs."),
    ("Outdated Training Data", C_MID,
     "Most AI models have a knowledge cutoff. Events, research, or policies after that "
     "date may not be reflected in responses."),
]
for i, (title, col, desc) in enumerate(reasons):
    row = i // 2; col_idx = i % 2
    bx = 0.3 + col_idx * 6.5; by = 1.55 + row * 2.8
    rect(sl, bx, by, 6.2, 2.6, col)
    txb(sl, title, bx + 0.2, by + 0.15, 5.8, 0.6,
        size=22, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.2), Inches(by + 0.8), Inches(5.8), Inches(1.65))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_LGRAY

# Example of hallucination
example_slide(
    "Hallucination Example",
    "Who wrote the novel 'The Language of Seasons' published in 2019?",
    "The novel 'The Language of Seasons' was written by Maria Chen and published by "
    "Penguin Random House in April 2019. It won the Booker Prize shortlist that year.\n\n"
    "⚠️  This book, author, and award are completely fabricated. "
    "Always verify book titles, authors, and awards with a library catalogue or publisher website.",
    subtitle="This is what a hallucination looks like — confident but false"
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 2  –  MINDSET SHIFT
# ═════════════════════════════════════════════════════════════════════════════
section_divider(2, "Mindset Shift: Search → Collaboration",
                "Change how you talk to AI and transform what you get back")

# Search mindset slide
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
    subtitle="Short, keyword‑style queries — passive information retrieval",
    bullet_color=C_RED,
    icon="🔎"
)

# Collaborator mindset slide
bullet_slide(
    "The 'Collaborator' Mindset",
    [
        '"Act as an environmental scientist. Explain climate change to a first‑year student in 3 bullet points."',
        '"Help me understand the main causes of WWI vs. WWII. Ask me questions first to check my knowledge gaps."',
        '"Give me a weekly meal plan as a table. I\'m vegetarian and on a budget."',
        '"I\'m writing an essay on leadership. Suggest 5 thesis angles and explain the pros of each."',
        '"I\'m a beginner. Teach me Python step‑by‑step, starting with variables. Check my understanding after each step."',
    ],
    subtitle="Specific, contextual, interactive — active collaboration",
    bullet_color=C_GREEN,
    icon="🤝"
)

# Side‑by‑side comparison
two_col_slide(
    "Search vs. Collaboration — Side by Side",
    "🔎  Search Mindset",
    [
        "Short, keyword queries",
        "Expects one correct answer",
        "Passive — reads and accepts",
        "No follow‑up or iteration",
        "Treats AI like a dictionary",
        "Frustrated when answer is vague",
    ],
    "🤝  Collaborator Mindset",
    [
        "Context‑rich, specific prompts",
        "Expects a starting point to refine",
        "Active — questions, pushes back",
        "Iterates and asks follow‑ups",
        "Treats AI like a smart assistant",
        "Uses vague answers as a chance to clarify",
    ],
    subtitle="Which mindset produces better results?"
)

# Pair discussion activity
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
    time_str="8 min"
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 3  –  PROMPTING LAB
# ═════════════════════════════════════════════════════════════════════════════
section_divider(3, "Prompting Lab",
                "4 Strategies  ·  Examples  ·  Try‑It Activities")

# ── Strategy 1: Role + Task ──────────────────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 1: Role + Task Prompting",
           subtitle="Tell AI WHO it is and WHAT you need it to do")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
txb(sl, "TEMPLATE", 0.5, 1.65, 3, 0.45,
    size=15, bold=True, color=C_TEAL)
txb(sl,
    '"Act as [ROLE]. Your audience is [AUDIENCE]. [TASK]. Use [TONE/FORMAT]."',
    0.5, 2.1, 12.3, 0.8,
    size=22, bold=True, color=C_YELLOW, italic=True)
rect(sl, 0.3, 3.0, 12.7, 0.06, C_ACCENT)
txb(sl, "Why it works:", 0.5, 3.1, 3, 0.4,
    size=17, bold=True, color=C_TEAL)
reasons = [
    "▸  Role sets the expertise level and vocabulary AI uses",
    "▸  Task gives a clear goal — no guessing",
    "▸  Audience shapes complexity and tone",
    "▸  Format ensures output is ready to use",
]
by = 3.55
for r in reasons:
    txb(sl, r, 0.6, by, 12.0, 0.45, size=19, color=C_LGRAY)
    by += 0.5

example_slide(
    "Strategy 1 — Role + Task Example",
    "Act as an academic English tutor. My student is a first‑year international student "
    "who struggles with thesis statements. Explain what a thesis statement is and give 3 "
    "examples — one weak, one okay, one strong. Use simple language.",
    "A thesis statement is the main argument of your essay — it tells your reader what "
    "you will prove.\n\n"
    "Weak: 'Social media is popular.'  (too vague — everyone agrees)\n"
    "Okay: 'Social media affects students.'  (needs to say HOW and WHY)\n"
    "Strong: 'Excessive social media use reduces academic focus because it fragments "
    "attention and displaces study time.' ✓",
    subtitle="Role: Tutor  |  Audience: First‑year student  |  Format: 3 examples"
)

activity_slide(
    "Try It: Role + Task",
    [
        "Write a Role + Task prompt for ONE of these situations:",
        "",
        "→  You are preparing for a job interview at a tech company.",
        "→  You need to explain your research topic to a non‑expert friend.",
        "→  You want feedback on an email you wrote to a professor.",
        "→  You are writing a cover letter for an internship.",
        "",
        "Use the template:",
        "→  'Act as [ROLE]. My audience is [AUDIENCE]. [TASK]. Use [FORMAT].'",
        "",
        "Share your prompt with a neighbour and compare results!",
    ],
    time_str="6 min", color=C_TEAL
)

# ── Strategy 2: Step‑by‑step / Scaffolded Prompting ─────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 2: Step‑by‑Step / Scaffolded Prompting",
           subtitle="Break complex tasks into guided stages")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
txb(sl, "HOW TO ASK FOR BREAKDOWNS", 0.5, 1.65, 8, 0.45,
    size=16, bold=True, color=C_TEAL)
steps = [
    ("1", "Start big", "Ask for an overview or outline first."),
    ("2", "Zoom in",   "Follow up: 'Now explain step 2 in more detail.'"),
    ("3", "Check in",  "Ask: 'What have I missed? What should I consider next?'"),
    ("4", "Iterate",   "Refine: 'Simplify this section. Add an example.'"),
]
for i, (num, title, desc) in enumerate(steps):
    bx = 0.4 + i * 3.2; by = 2.2
    rect(sl, bx, by, 3.0, 4.7, C_ACCENT)
    txb(sl, num, bx + 0.1, by + 0.1, 0.8, 0.8,
        size=36, bold=True, color=C_TEAL)
    txb(sl, title, bx + 0.1, by + 1.0, 2.8, 0.55,
        size=20, bold=True, color=C_YELLOW)
    box = sl.shapes.add_textbox(Inches(bx + 0.1), Inches(by + 1.65), Inches(2.8), Inches(2.9))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_LGRAY

example_slide(
    "Strategy 2 — Scaffolded Prompting Example",
    "I need to write a 1,500‑word essay on globalisation and education. "
    "Start by giving me an outline with 5 sections and 2–3 bullet points per section. "
    "Do not write the essay yet — just the structure.",
    "Outline: Globalisation & Education\n"
    "1. Introduction  •  Define globalisation  •  Relevance to education  •  Thesis preview\n"
    "2. Benefits  •  Access to resources  •  Cultural exchange  •  International opportunities\n"
    "3. Challenges  •  Inequality of access  •  Language barriers  •  Cultural homogenisation\n"
    "4. Case Studies  •  Online learning platforms  •  International student mobility\n"
    "5. Conclusion  •  Summary  •  Future implications  •  Call to action",
    subtitle="Notice: you control the pace, AI builds step by step"
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
    time_str="7 min", color=C_GREEN
)

# ── Strategy 3: Format‑based Output ──────────────────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 3: Format‑Based Output",
           subtitle="Tell AI exactly how you want the answer structured")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
formats = [
    ("Tables",     "Compare options side by side.\n'Give me a table comparing X, Y, Z by cost, features, and ease of use.'"),
    ("Checklists", "Get actionable to‑do lists.\n'Give me a checklist of things to do before submitting my assignment.'"),
    ("Templates",  "Generate reusable structures.\n'Create an email template I can use to ask professors for extensions.'"),
    ("Bullet Points","Scannable summaries.\n'Summarise this article in 5 bullet points, each under 15 words.'"),
    ("Numbered Lists","Step‑by‑step instructions.\n'Give me numbered steps to set up a Python environment.'"),
    ("Paragraphs", "Flowing prose for essays.\n'Write this as 2 short paragraphs suitable for an academic essay.'"),
]
for i, (fmt, desc) in enumerate(formats):
    row = i // 3; col = i % 3
    bx = 0.3 + col * 4.35; by = 1.6 + row * 2.75
    rect(sl, bx, by, 4.1, 2.55, C_ACCENT if row == 0 else C_MID)
    txb(sl, fmt, bx + 0.15, by + 0.12, 3.8, 0.55,
        size=20, bold=True, color=C_TEAL)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.15), Inches(by + 0.72), Inches(3.8), Inches(1.72))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(15); run.font.color.rgb = C_LGRAY

example_slide(
    "Strategy 3 — Format‑Based Output Example",
    "Compare three productivity apps (Notion, Todoist, Trello) using a table. "
    "Columns: App Name, Best For, Cost, Ease of Use (1–5 stars), Mobile App. "
    "Add a recommendation row at the bottom.",
    "| App      | Best For         | Cost        | Ease ⭐ | Mobile |\n"
    "|----------|-----------------|-------------|--------|--------|\n"
    "| Notion   | Notes + projects | Free / $8/mo |  ★★★★  |  ✓   |\n"
    "| Todoist  | Task management  | Free / $4/mo |  ★★★★★ |  ✓   |\n"
    "| Trello   | Visual boards    | Free / $5/mo |  ★★★★  |  ✓   |\n"
    "| 🏆 Pick  | Beginners → Todoist  |  Teams → Trello  |  Power users → Notion |",
    subtitle="Formatted output is immediately usable — no reformatting required"
)

activity_slide(
    "Try It: Format‑Based Output",
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
        "→  'Give me a checklist of 8 steps for managing exam stress, with a "
        "☐ checkbox before each item.'",
        "",
        "Compare outputs with a partner who chose the same topic but a different format!",
    ],
    time_str="6 min", color=C_ORANGE
)

# ── Strategy 4: Revision & Reflection Prompting ───────────────────────────────
sl = add_slide()
header_bar(sl, "Strategy 4: Revision & Reflection Prompting",
           subtitle="Use AI as an editor, critic, and thinking partner")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
techniques = [
    ("Ask for Improvement", C_ACCENT,
     "Paste your draft and say:\n'Improve the clarity and flow of this paragraph. "
     "Keep my original voice.'"),
    ("Ask for Questions First", C_MID,
     "Before a big task say:\n'Before you help me, ask me 5 questions to better "
     "understand my needs and context.'"),
    ("Ask for Critique", RGBColor(0x2D, 0x00, 0x5E),
     "After a draft say:\n'What are the three weakest parts of this argument? "
     "How would you strengthen each one?'"),
    ("Ask for Alternatives", C_ACCENT,
     "After an output say:\n'Give me 3 alternative ways to phrase this conclusion. "
     "Rate each one and explain why.'"),
]
for i, (title, col, desc) in enumerate(techniques):
    bx = 0.3 + (i % 2) * 6.5; by = 1.6 + (i // 2) * 2.75
    rect(sl, bx, by, 6.2, 2.55, col)
    txb(sl, title, bx + 0.2, by + 0.12, 5.8, 0.55,
        size=21, bold=True, color=C_TEAL)
    box = sl.shapes.add_textbox(
        Inches(bx + 0.2), Inches(by + 0.75), Inches(5.8), Inches(1.65))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_LGRAY

example_slide(
    "Strategy 4 — Revision Example",
    "Here is my conclusion paragraph: [paste paragraph]. "
    "Identify 2 weaknesses and suggest specific improvements. "
    "Then give me a revised version that is more academic in tone.",
    "Weaknesses identified:\n"
    "1. Your conclusion repeats the introduction almost word‑for‑word — "
    "it should synthesise rather than summarise.\n"
    "2. The final sentence is vague: 'AI is important for the future.' "
    "— add a specific implication.\n\n"
    "Revised conclusion:\n"
    "'Taken together, these findings suggest that AI literacy is no longer optional "
    "for university students — it is a foundational skill for participation in the "
    "modern knowledge economy. Institutions that integrate AI training into curricula "
    "will produce graduates better equipped for the challenges ahead.'",
    subtitle="AI as editor — keeps your voice, strengthens your argument"
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
    time_str="8 min", color=C_TEAL
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 4  –  EVALUATING AI OUTPUT
# ═════════════════════════════════════════════════════════════════════════════
section_divider(4, "Evaluating AI Output",
                "Not all AI answers are equal — learn to tell the difference")

# Why different tools give different answers
sl = add_slide()
header_bar(sl, "Why Different AI Tools Give Different Answers",
           subtitle="Same question, different outputs — here's why")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
reasons = [
    ("Different Training Data",
     "Each AI model is trained on different datasets with different time cutoffs, "
     "languages, and sources of information."),
    ("Different Model Architecture",
     "GPT‑4, Gemini, Claude, and Llama use different underlying approaches to "
     "language modelling — producing different styles and strengths."),
    ("Different System Prompts",
     "Every AI product has hidden instructions that shape tone, safety limits, "
     "and focus areas — you don't always see these."),
    ("Randomness (Temperature)",
     "AI responses include controlled randomness. Run the same prompt twice "
     "and you may get different outputs."),
    ("Context Length",
     "Some tools handle longer conversations or documents better than others, "
     "affecting the quality of responses in extended tasks."),
]
by = 1.7
for r_title, r_desc in reasons:
    rect(sl, 0.4, by, 12.4, 1.0, C_ACCENT)
    txb(sl, r_title, 0.6, by + 0.08, 4.5, 0.45,
        size=18, bold=True, color=C_TEAL)
    box = sl.shapes.add_textbox(Inches(5.2), Inches(by + 0.08), Inches(7.5), Inches(0.82))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = r_desc; run.font.size = Pt(16); run.font.color.rgb = C_LGRAY
    by += 1.07

# Evaluation checklist
sl = add_slide()
header_bar(sl, "AI Output Evaluation Checklist",
           subtitle="Use this every time before you trust or submit AI content")
checklist = [
    ("☐  Accuracy",          C_RED,    "Can you verify key facts from a reliable primary source?"),
    ("☐  Bias",              C_ORANGE, "Does the response favour one perspective unfairly?"),
    ("☐  Missing Citations", C_YELLOW, "Are sources named? Are they real and accessible?"),
    ("☐  Clarity",           C_TEAL,   "Is the language clear, relevant, and appropriate for your audience?"),
    ("☐  Cultural Relevance",C_GREEN,  "Does it account for your cultural or regional context?"),
    ("☐  Completeness",      C_ACCENT, "Does it fully address your question, or are there gaps?"),
]
for i, (item, col, desc) in enumerate(checklist):
    by = 1.55 + i * 0.97
    rect(sl, 0.3, by, 12.7, 0.88, col)
    txb(sl, item, 0.5, by + 0.12, 4.2, 0.65,
        size=20, bold=True, color=C_WHITE)
    box = sl.shapes.add_textbox(Inches(4.9), Inches(by + 0.12), Inches(8.0), Inches(0.65))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = desc; run.font.size = Pt(17); run.font.color.rgb = C_DARK; run.font.bold = False

# Practice example with errors
sl = add_slide()
header_bar(sl, "Practice: Spot the Errors",
           subtitle="Read this AI output — how many problems can you find?",
           bar_color=C_RED, title_color=C_WHITE, sub_color=C_YELLOW)
rect(sl, 0.3, 1.55, 12.7, 4.1, C_MID)
txb(sl, "AI OUTPUT (contains deliberate errors):", 0.5, 1.65, 9, 0.45,
    size=15, bold=True, color=C_RED)
error_text = (
    "The United Nations was founded in 1945 by Franklin D. Roosevelt and Winston Churchill "
    "at the Geneva Conference. It currently has 184 member states. The UN's primary goal "
    "is to prevent wars and promote peace worldwide. Its headquarters are located in Paris, "
    "France. The Secretary-General serves a 3-year term. The UN Security Council has "
    "7 permanent members including the USA, UK, France, Russia, and China."
)
box = sl.shapes.add_textbox(Inches(0.5), Inches(2.15), Inches(12.2), Inches(3.3))
tf = box.text_frame; tf.word_wrap = True
run = tf.paragraphs[0].add_run()
run.text = error_text; run.font.size = Pt(19); run.font.color.rgb = C_LGRAY; run.font.italic = True

rect(sl, 0.3, 5.75, 12.7, 1.45, RGBColor(0x2D, 0x00, 0x5E))
txb(sl, "❓  Errors to find: (reveal on click / discuss with partner)",
    0.5, 5.82, 12.0, 0.4, size=16, bold=True, color=C_YELLOW)
answers = ("Errors: ①  Geneva → San Francisco  "
           "②  184 → 193 members  "
           "③  HQ is New York, not Paris  "
           "④  Term is 5 years, not 3  "
           "⑤  Security Council has 5 permanent members, not 7")
txb(sl, answers, 0.5, 6.28, 12.5, 0.8, size=15, color=C_GREEN)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 5  –  AI GUIDELINES / RESPONSIBLE USE
# ═════════════════════════════════════════════════════════════════════════════
section_divider(5, "AI Guidelines & Responsible Use",
                "Know the rules — protect yourself and others")

# When AI is allowed
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
    bullet_color=C_GREEN,
    icon="✓"
)

# When AI is NOT allowed
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
    bullet_color=C_RED,
    icon="✗"
)

# Privacy + safety
sl = add_slide()
header_bar(sl, "Privacy, Safety & Personal Data",
           subtitle="What you share with AI is not always private")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
privacy_items = [
    ("🔒 Never Share", C_RED, [
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
    bx = 0.4 + i * 6.5
    rect(sl, bx, 1.6, 6.2, 5.5, col)
    txb(sl, heading, bx + 0.2, 1.72, 5.8, 0.65,
        size=22, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    box = sl.shapes.add_textbox(Inches(bx + 0.2), Inches(2.5), Inches(5.8), Inches(4.4))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for item in items:
        if first: p = tf.paragraphs[0]; first = False
        else: p = tf.add_paragraph()
        p.space_before = Pt(9)
        run = p.add_run(); run.text = f"▸  {item}"
        run.font.size = Pt(18); run.font.color.rgb = C_WHITE

# Ethical use
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
    subtitle="The difference is whether YOU are growing"
)

# Citation expectations
bullet_slide(
    "Basic AI Citation Expectations",
    [
        "Check your institution's style guide — many now include AI citation formats.",
        "APA 7 example:  OpenAI. (2024). ChatGPT (GPT‑4) [Large language model]. https://chat.openai.com",
        "MLA example:  'Prompt text.' ChatGPT, GPT‑4 version, OpenAI, 15 Jan. 2024, chat.openai.com.",
        "Include: AI tool name  ·  Version  ·  Date of conversation  ·  Prompt used (if required)",
        "When in doubt, be transparent — disclose AI use in a brief author's note.",
        "Remember: citations show intellectual honesty, not weakness.",
    ],
    subtitle="When disclosure is required, cite clearly and completely",
    bullet_color=C_LGRAY,
    icon="📎"
)

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 6  –  FINAL CHALLENGE + REFLECTION
# ═════════════════════════════════════════════════════════════════════════════
section_divider(6, "Final Prompting Challenge + Reflection",
                "Put all 4 strategies together in one powerful prompt")

# Build-your-best-prompt
sl = add_slide()
header_bar(sl, "Build Your Best Prompt Challenge",
           subtitle="Combine all 6 elements into one masterful prompt",
           bar_color=C_ORANGE, title_color=C_DARK, sub_color=C_DARK)
elements = [
    ("1", "Role",          "Who should AI be?",                       "Act as a …"),
    ("2", "Context",       "What's the background?",                  "I am a … working on …"),
    ("3", "Task",          "What exactly do you need?",               "Help me … by …"),
    ("4", "Constraints",   "What limits or requirements apply?",      "Keep it under … / Do not …"),
    ("5", "Output Format", "How should the answer look?",             "Format as a table / checklist / …"),
    ("6", "Language Level","What reading level / language?",          "Use B2 English / avoid jargon / …"),
]
for i, (num, label, question, hint) in enumerate(elements):
    row = i // 3; col = i % 3
    bx = 0.25 + col * 4.35; by = 1.55 + row * 2.8
    rect(sl, bx, by, 4.1, 2.6, C_ACCENT if row == 0 else C_MID)
    txb(sl, num, bx + 0.1, by + 0.1, 0.7, 0.65,
        size=28, bold=True, color=C_TEAL)
    txb(sl, label, bx + 0.8, by + 0.1, 3.1, 0.6,
        size=20, bold=True, color=C_YELLOW)
    txb(sl, question, bx + 0.1, by + 0.82, 3.9, 0.55,
        size=15, color=C_LGRAY, italic=True)
    txb(sl, hint, bx + 0.1, by + 1.45, 3.9, 0.9,
        size=14, color=C_TEAL)

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

# Reflection prompts
sl = add_slide()
header_bar(sl, "Reflection: Mindset Shift & Next Steps",
           subtitle="Take 3 minutes to think — then share one insight")
rect(sl, 0.3, 1.55, 12.7, 5.6, C_MID)
reflections = [
    ("🔄", "Mindset Shift",
     "At the start of today, how did you think about AI? How has that changed?"),
    ("💡", "Biggest Insight",
     "What is the single most useful thing you learned in this workshop?"),
    ("✏️", "First Prompt",
     "What is the first real prompt you will try using today's strategies?"),
    ("⚠️", "Responsible Use",
     "What is one way you will use AI MORE responsibly going forward?"),
    ("🚀", "Next Step",
     "What skill do you want to develop further — prompting, evaluation, or something else?"),
]
by = 1.65
for emoji, title, question in reflections:
    rect(sl, 0.4, by, 12.4, 1.0, C_ACCENT)
    txb(sl, f"{emoji}  {title}", 0.6, by + 0.08, 3.5, 0.5,
        size=18, bold=True, color=C_TEAL)
    box = sl.shapes.add_textbox(Inches(4.2), Inches(by + 0.1), Inches(8.4), Inches(0.78))
    tf = box.text_frame; tf.word_wrap = True
    run = tf.paragraphs[0].add_run()
    run.text = question; run.font.size = Pt(17); run.font.color.rgb = C_LGRAY
    by += 1.06

# ═════════════════════════════════════════════════════════════════════════════
# SECTION 7  –  CLOSING SLIDE
# ═════════════════════════════════════════════════════════════════════════════
section_divider(7, "Closing & Resources",
                "Thank you — now go collaborate!")

sl = add_slide()
rect(sl, 0, 0, 0.5, 7.5, C_GREEN)
rect(sl, 0.5, 0, 0.15, 7.5, C_TEAL)
txb(sl, "Thank You! 🎉", 1.0, 0.4, 11.5, 1.0,
    size=44, bold=True, color=C_WHITE)
txb(sl, "You are now an AI collaborator — not just an AI user.",
    1.0, 1.5, 11.5, 0.65, size=24, color=C_TEAL, italic=True)

# Resources box
rect(sl, 0.8, 2.3, 11.7, 4.8, C_MID)
txb(sl, "Resources & Prompt Templates", 1.0, 2.4, 11.0, 0.6,
    size=22, bold=True, color=C_YELLOW)
resources = [
    "🔗  ChatGPT:  chat.openai.com         🔗  Claude:  claude.ai",
    "🔗  Gemini:   gemini.google.com        🔗  Copilot: copilot.microsoft.com",
    "",
    "📋  Master Prompt Template:",
    "     'Act as [ROLE]. I am [CONTEXT]. Help me [TASK]. Keep [CONSTRAINTS]. Format as [FORMAT]. Use [LANGUAGE].'",
    "",
    "📚  Further Learning:",
    "     • Learn Prompting (free):  learnprompting.org",
    "     • AI Literacy (free):      elementsofai.com",
    "     • Prompt Engineering Guide: promptingguide.ai",
    "",
    "📧  Workshop Support:  [facilitator@institution.edu]",
    "📞  Student Services:  [Your institution's student support contact]",
]
box = sl.shapes.add_textbox(Inches(1.0), Inches(3.1), Inches(11.2), Inches(3.8))
tf = box.text_frame; tf.word_wrap = True
first = True
for line in resources:
    if first: p = tf.paragraphs[0]; first = False
    else: p = tf.add_paragraph()
    p.space_before = Pt(5)
    run = p.add_run(); run.text = line
    run.font.size = Pt(16)
    if line.startswith("📋") or line.startswith("📚"):
        run.font.color.rgb = C_TEAL; run.font.bold = True
    elif line.startswith("     'Act"):
        run.font.color.rgb = C_YELLOW; run.font.italic = True
    else:
        run.font.color.rgb = C_LGRAY

txb(sl, "Remember: The best prompt is the one that helps you think more clearly.",
    0.8, 7.05, 11.7, 0.4, size=15, color=C_GREEN, italic=True,
    align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────────────────────────────────────
out = "/home/user/powerpoints/AI_Workshop_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}  ({len(prs.slides)} slides)")
