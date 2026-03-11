from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Color Palette ─────────────────────────────────────────────────────────────
C_BG      = RGBColor(0xFF, 0xFF, 0xFF)
C_BG2     = RGBColor(0xF4, 0xF6, 0xFF)
C_PURPLE  = RGBColor(0x6C, 0x35, 0xDE)
C_BLUE    = RGBColor(0x22, 0x8B, 0xE6)
C_TEAL    = RGBColor(0x00, 0xBF, 0xA5)
C_GREEN   = RGBColor(0x00, 0xC8, 0x53)
C_ORANGE  = RGBColor(0xFF, 0x6D, 0x00)
C_PINK    = RGBColor(0xE9, 0x1E, 0x8C)
C_YELLOW  = RGBColor(0xFF, 0xD6, 0x00)
C_DARK    = RGBColor(0x1A, 0x1A, 0x2E)
C_DGRAY   = RGBColor(0x44, 0x44, 0x66)
C_LGRAY   = RGBColor(0xE8, 0xEC, 0xF8)
C_WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
C_LEAF    = RGBColor(0x2E, 0x7D, 0x32)
C_DGREEN  = RGBColor(0x1B, 0x5E, 0x20)

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

BLANK = prs.slide_layouts[6]

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
    run.font.size      = Pt(size)
    run.font.bold      = bold
    run.font.color.rgb = color
    run.font.italic    = italic
    return run


def multi_txb(sl, lines, x, y, w, h,
              size=20, color=C_DARK, align=PP_ALIGN.LEFT,
              bold=False, spacing=8, icon=""):
    """Add a text box with multiple paragraphs."""
    box = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]; first = False
        else:
            p = tf.add_paragraph()
        p.space_before = Pt(spacing)
        p.alignment = align
        run = p.add_run()
        run.text = f"{icon}  {line}" if icon else line
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.bold = bold
    return tf


def rect(sl, x, y, w, h, fill_color):
    shape = sl.shapes.add_shape(
        1, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    return shape


def rounded_rect(sl, x, y, w, h, fill_color):
    shape = sl.shapes.add_shape(
        5, Inches(x), Inches(y), Inches(w), Inches(h))
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
    txb(sl, str(section_num), 8.5, 0.5, 4.5, 6.5,
        size=280, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
        align=PP_ALIGN.LEFT)
    rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
    txb(sl, f"SECTION {section_num}", 0.85, 2.1, 10, 0.55,
        size=18, bold=True, color=C_YELLOW, italic=True)
    txb(sl, section_title, 0.85, 2.75, 10.5, 1.5,
        size=46, bold=True, color=C_WHITE)
    if subtitle:
        txb(sl, subtitle, 0.85, 4.4, 10.5, 0.85,
            size=22, color=C_LGRAY, italic=True)
    return sl


def framework_card(sl, letter, label, description, x, y, w, h,
                   letter_color=C_PURPLE, bg_color=C_BG2):
    """Draw a rounded card for one letter of a framework acronym."""
    rounded_rect(sl, x, y, w, h, bg_color)
    txb(sl, letter, x + 0.15, y + 0.08, 0.6, 0.55,
        size=30, bold=True, color=letter_color)
    txb(sl, f"= {label}", x + 0.7, y + 0.12, w - 0.9, 0.45,
        size=20, bold=True, color=C_DARK)
    txb(sl, description, x + 0.25, y + 0.55, w - 0.45, h - 0.65,
        size=16, color=C_DGRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1: TITLE SLIDE
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide(bg=C_PURPLE)
rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
txb(sl, "INTERMEDIATE", 0.85, 1.4, 12, 0.8,
    size=22, bold=True, color=C_YELLOW, italic=True)
txb(sl, "AI Prompting", 0.85, 2.1, 12, 1.3,
    size=56, bold=True, color=C_WHITE)
txb(sl, "Frameworks & Techniques", 0.85, 3.2, 12, 0.8,
    size=38, bold=False, color=C_LGRAY)
rect(sl, 0.85, 4.2, 4.5, 0.06, C_YELLOW)
txb(sl, "From Good Prompts to Great Results", 0.85, 4.6, 12, 0.7,
    size=22, color=C_LGRAY, italic=True)
txb(sl, "Beyond the basics: structured frameworks, green prompting,\n"
        "and critical thinking for reliable AI outputs.",
    0.85, 5.4, 10, 1.0,
    size=18, color=C_LGRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2: WHY GOOD PROMPTING MATTERS
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Why Does Prompting Matter?",
           subtitle="Better prompts = better results, faster")

# Left column - benefits
rounded_rect(sl, 0.4, 1.65, 6.0, 5.5, C_BG2)
txb(sl, "Better Prompts Help You...", 0.6, 1.75, 5.6, 0.5,
    size=22, bold=True, color=C_PURPLE)

benefits = [
    ("Get better results, faster",
     "Reduce back-and-forth by being specific upfront"),
    ("Reduce ambiguity",
     "Clear instructions leave less room for misinterpretation"),
    ("Guide LLM reasoning paths",
     "Structure steers the AI toward the logic you need"),
    ("Mitigate harmful or biased outputs",
     "Guardrails in your prompt shape safer responses"),
    ("Improve consistency & reliability",
     "Repeatable frameworks yield repeatable quality"),
]
y_pos = 2.35
for title, desc in benefits:
    txb(sl, f"  {title}", 0.6, y_pos, 5.6, 0.35,
        size=18, bold=True, color=C_DARK)
    txb(sl, f"     {desc}", 0.6, y_pos + 0.32, 5.6, 0.35,
        size=14, color=C_DGRAY)
    y_pos += 0.82

# Right column - critical thinking
rounded_rect(sl, 6.8, 1.65, 6.1, 5.5, RGBColor(0xFD, 0xF0, 0xE7))
txb(sl, "Stay in the Driver's Seat", 6.95, 1.75, 5.8, 0.5,
    size=22, bold=True, color=C_ORANGE)
txb(sl, "The more context and structure YOU provide,\n"
        "the more YOU stay in charge of the content.",
    7.0, 2.45, 5.5, 1.0,
    size=18, bold=True, color=C_DARK)
txb(sl, "Good prompting is critical thinking in action.\n\n"
        "When you invest effort into crafting a prompt, you are:\n\n"
        "  Clarifying your own goals and requirements\n\n"
        "  Thinking through what \"good\" looks like\n\n"
        "  Evaluating outputs against your standards\n\n"
        "  Keeping human judgment at the center\n\n"
        "AI is a tool — your expertise drives the outcome.",
    7.0, 3.4, 5.5, 4.0,
    size=16, color=C_DGRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION DIVIDER: PROMPTING FRAMEWORKS
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(1, "Prompting Frameworks",
                subtitle="Structured approaches for consistent, high-quality prompts")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3: RTF FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "RTF Framework", subtitle="Role — Task — Format", bar_color=C_BLUE)

framework_card(sl, "R", "Role",
               "Who should the AI act as?\n"
               "e.g., \"You are a senior data analyst...\"",
               0.4, 1.7, 3.8, 1.65, letter_color=C_BLUE)
framework_card(sl, "T", "Task",
               "What specific task should it perform?\n"
               "e.g., \"Analyze Q3 sales trends and identify...\"",
               0.4, 3.55, 3.8, 1.65, letter_color=C_BLUE)
framework_card(sl, "F", "Format",
               "How should the output be structured?\n"
               "e.g., \"Present as a bulleted summary with...\"",
               0.4, 5.4, 3.8, 1.65, letter_color=C_BLUE)

# Example panel
rounded_rect(sl, 4.6, 1.7, 8.3, 5.35, RGBColor(0xE8, 0xF4, 0xFD))
txb(sl, "Example Prompt Using RTF", 4.8, 1.8, 7.9, 0.45,
    size=20, bold=True, color=C_BLUE)
txb(sl, "[Role] You are an experienced marketing strategist\n"
        "who specializes in social media campaigns for\n"
        "small businesses.\n\n"
        "[Task] Create a 1-week social media content\n"
        "calendar for a local bakery launching a new\n"
        "line of gluten-free pastries. Include post ideas\n"
        "for Instagram, Facebook, and TikTok.\n\n"
        "[Format] Present the calendar as a table with\n"
        "columns for Day, Platform, Post Type, Caption\n"
        "Idea, and Hashtag Suggestions.",
    4.8, 2.35, 7.9, 4.5,
    size=16, color=C_DARK)

txb(sl, "Best for: Quick, everyday prompts that need a clear structure",
    0.4, 7.15, 12.5, 0.35,
    size=14, bold=True, color=C_BLUE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4: CREATE FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "CREATE Framework",
           subtitle="Character — Request — Examples — Adjustments — Type of Output — Extras",
           bar_color=C_TEAL)

cards = [
    ("C", "Character", "Define the AI's persona or expertise"),
    ("R", "Request", "State the core task or question clearly"),
    ("E", "Examples", "Provide sample inputs/outputs to guide style"),
    ("A", "Adjustments", "Specify constraints, tone, or limitations"),
    ("T", "Type of Output", "Define the desired format (list, essay, code, etc.)"),
    ("E", "Extras", "Add any additional context or special instructions"),
]

col1_x, col2_x = 0.4, 6.9
y_start = 1.7
card_h = 1.55

for i, (letter, label, desc) in enumerate(cards):
    x = col1_x if i < 3 else col2_x
    y = y_start + (i % 3) * (card_h + 0.2)
    framework_card(sl, letter, label, desc,
                   x, y, 6.0, card_h, letter_color=C_TEAL)

txb(sl, "Best for: Detailed, nuanced prompts where tone, examples, and constraints matter",
    0.4, 7.15, 12.5, 0.35,
    size=14, bold=True, color=C_TEAL, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5: CO-STAR FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "CO-STAR Framework",
           subtitle="Context — Objective — Style — Tone — Audience — Response",
           bar_color=C_PURPLE)

costar = [
    ("C", "Context", "Background info & situation the AI needs to know"),
    ("O", "Objective", "The specific goal or task you want accomplished"),
    ("S", "Style", "Writing style — academic, casual, technical, etc."),
    ("T", "Tone", "Emotional quality — encouraging, formal, witty, etc."),
    ("A", "Audience", "Who will consume the output? Tailor appropriately"),
    ("R", "Response", "Desired output format — paragraph, list, JSON, etc."),
]

for i, (letter, label, desc) in enumerate(costar):
    x = col1_x if i < 3 else col2_x
    y = y_start + (i % 3) * (card_h + 0.2)
    framework_card(sl, letter, label, desc,
                   x, y, 6.0, card_h, letter_color=C_PURPLE)

txb(sl, "Best for: Content creation tasks — blogs, emails, reports where style and audience matter",
    0.4, 7.15, 12.5, 0.35,
    size=14, bold=True, color=C_PURPLE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6: CRISPE FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "CRISPE Framework",
           subtitle="Capacity — Request — Insight — Statement — Personality — Experiment",
           bar_color=C_PINK)

crispe = [
    ("C", "Capacity & Role", "What role or capacity should the AI assume?"),
    ("R", "Request", "What specific output are you asking for?"),
    ("I", "Insight", "What background knowledge or data should it consider?"),
    ("S", "Statement", "Clear statement of what you want delivered"),
    ("P", "Personality", "What style or personality should the response reflect?"),
    ("E", "Experiment", "Ask for multiple variations or options to compare"),
]

for i, (letter, label, desc) in enumerate(crispe):
    x = col1_x if i < 3 else col2_x
    y = y_start + (i % 3) * (card_h + 0.2)
    framework_card(sl, letter, label, desc,
                   x, y, 6.0, card_h, letter_color=C_PINK)

txb(sl, "Best for: Creative and exploratory tasks — brainstorming, ideation, generating multiple options",
    0.4, 7.15, 12.5, 0.35,
    size=14, bold=True, color=C_PINK, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7: RISEN FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "RISEN Framework",
           subtitle="Role — Instructions — Steps — End Goal — Narrowing",
           bar_color=C_ORANGE)

risen = [
    ("R", "Role", "Assign a specific role or expert persona"),
    ("I", "Instructions", "Provide clear, detailed instructions"),
    ("S", "Steps", "Break the task into sequential steps"),
    ("E", "End Goal", "Define what success looks like"),
    ("N", "Narrowing", "Add constraints to focus the output"),
]

# Use single column layout for 5 items
card_h_small = 1.0
for i, (letter, label, desc) in enumerate(risen):
    x = 0.4 if i < 3 else 6.9
    y = y_start + (i % 3) * (card_h_small + 0.15)
    framework_card(sl, letter, label, desc,
                   x, y, 6.0, card_h_small, letter_color=C_ORANGE)

# Example panel for RISEN
rounded_rect(sl, 0.4, 5.2, 12.5, 1.8, RGBColor(0xFD, 0xF0, 0xE7))
txb(sl, "When to use RISEN:", 0.6, 5.3, 5.0, 0.35,
    size=18, bold=True, color=C_ORANGE)
txb(sl, "Ideal for multi-step tasks where you need procedural accuracy — like writing a project plan, "
        "debugging code, or building a lesson plan. The \"Steps\" and \"Narrowing\" components make it "
        "especially powerful for complex, structured outputs.",
    0.6, 5.7, 12.0, 1.2,
    size=16, color=C_DGRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8: RACE FRAMEWORK
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "RACE Framework",
           subtitle="Role — Action — Context — Expectation",
           bar_color=C_GREEN)

race = [
    ("R", "Role", "Define who the AI should be\ne.g., \"You are a UX researcher...\""),
    ("A", "Action", "Specify the action to perform\ne.g., \"Conduct a heuristic evaluation...\""),
    ("C", "Context", "Provide relevant background\ne.g., \"...of our mobile checkout flow for Gen Z users\""),
    ("E", "Expectation", "State the expected deliverable\ne.g., \"Provide a ranked list of 10 usability issues...\""),
]

for i, (letter, label, desc) in enumerate(race):
    x = 0.4 if i < 2 else 6.9
    y = y_start + (i % 2) * 2.0
    framework_card(sl, letter, label, desc,
                   x, y, 6.0, 1.8, letter_color=C_GREEN)

# Comparison note
rounded_rect(sl, 0.4, 5.8, 12.5, 1.35, RGBColor(0xE8, 0xF5, 0xE9))
txb(sl, "RTF vs RACE:", 0.6, 5.9, 3.0, 0.35,
    size=18, bold=True, color=C_GREEN)
txb(sl, "RACE builds on RTF by adding explicit Context and setting clear Expectations for the deliverable. "
        "Use RTF for quick prompts; upgrade to RACE when you need more precision about the situation and outcome.",
    0.6, 6.25, 12.0, 0.85,
    size=16, color=C_DGRAY)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9: FRAMEWORK COMPARISON
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Choosing the Right Framework",
           subtitle="Match the framework to your task")

# Table header
rect(sl, 0.4, 1.65, 12.5, 0.55, C_PURPLE)
cols = [("Framework", 0.4, 2.0), ("Components", 2.6, 3.5),
        ("Best For", 6.3, 3.5), ("Complexity", 10.0, 2.5)]
for label, x, w in cols:
    txb(sl, label, x, 1.68, w, 0.45,
        size=16, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

# Table rows
frameworks = [
    ("RTF", "Role, Task, Format", "Quick everyday prompts", "Low"),
    ("CREATE", "Character, Request, Examples,\nAdjustments, Type, Extras", "Detailed, nuanced tasks", "High"),
    ("CO-STAR", "Context, Objective, Style,\nTone, Audience, Response", "Content creation & writing", "Medium"),
    ("CRISPE", "Capacity, Request, Insight,\nStatement, Personality, Experiment", "Creative & exploratory work", "Medium"),
    ("RISEN", "Role, Instructions, Steps,\nEnd Goal, Narrowing", "Multi-step procedural tasks", "Medium"),
    ("RACE", "Role, Action, Context,\nExpectation", "Precise deliverable-focused tasks", "Low-Med"),
]

row_colors = [C_BG2, C_WHITE]
row_h = 0.85
for i, (name, components, best_for, complexity) in enumerate(frameworks):
    y = 2.25 + i * row_h
    rect(sl, 0.4, y, 12.5, row_h, row_colors[i % 2])
    txb(sl, name, 0.55, y + 0.08, 1.8, row_h - 0.1,
        size=16, bold=True, color=C_PURPLE)
    txb(sl, components, 2.6, y + 0.05, 3.5, row_h - 0.1,
        size=13, color=C_DARK)
    txb(sl, best_for, 6.3, y + 0.08, 3.5, row_h - 0.1,
        size=14, color=C_DARK)
    txb(sl, complexity, 10.0, y + 0.08, 2.5, row_h - 0.1,
        size=14, color=C_DGRAY, align=PP_ALIGN.CENTER)

txb(sl, "Tip: You don't have to use a framework rigidly — mix and match elements that fit your task!",
    0.4, 7.15, 12.5, 0.35,
    size=14, bold=True, color=C_PURPLE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10: INTERACTIVE ACTIVITY 1 — TRY RTF
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Interactive Activity: Try the RTF Framework",
           subtitle="Your turn! Build a prompt using Role — Task — Format",
           bar_color=C_ORANGE)

rounded_rect(sl, 0.4, 1.65, 6.0, 5.5, RGBColor(0xFD, 0xF0, 0xE7))
txb(sl, "Your Challenge", 0.6, 1.75, 5.6, 0.45,
    size=22, bold=True, color=C_ORANGE)

txb(sl, "Pick one of these scenarios (or choose your own!):\n\n"
        "1. You need help writing a professional email\n"
        "   to reschedule a meeting with a client\n\n"
        "2. You want a study guide for a topic\n"
        "   you're learning\n\n"
        "3. You need a meal prep plan for the week\n\n"
        "Now build your prompt using RTF:\n\n"
        "  R  — Who should the AI be?\n"
        "  T  — What should it do?\n"
        "  F  — How should the answer look?",
    0.6, 2.3, 5.6, 4.5,
    size=16, color=C_DARK)

rounded_rect(sl, 6.8, 1.65, 6.1, 5.5, C_BG2)
txb(sl, "Write Your Prompt Here", 7.0, 1.75, 5.7, 0.45,
    size=22, bold=True, color=C_PURPLE)

txb(sl, "Role:\n"
        "_________________________________\n"
        "_________________________________\n\n"
        "Task:\n"
        "_________________________________\n"
        "_________________________________\n"
        "_________________________________\n\n"
        "Format:\n"
        "_________________________________\n"
        "_________________________________\n\n\n"
        "Then test it with your favorite AI tool!",
    7.0, 2.4, 5.7, 4.5,
    size=16, color=C_DGRAY)

txb(sl, "Time: 5 minutes  |  Try it out and share your results!",
    0.4, 7.15, 12.5, 0.35,
    size=16, bold=True, color=C_ORANGE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11: INTERACTIVE ACTIVITY 2 — TRY CO-STAR
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Interactive Activity: Try the CO-STAR Framework",
           subtitle="Apply CO-STAR to a content creation challenge",
           bar_color=C_ORANGE)

rounded_rect(sl, 0.4, 1.65, 6.0, 5.5, RGBColor(0xFD, 0xF0, 0xE7))
txb(sl, "Your Challenge", 0.6, 1.75, 5.6, 0.45,
    size=22, bold=True, color=C_ORANGE)

txb(sl, "Scenario: You need to write a LinkedIn post\n"
        "announcing a new product, service, or initiative\n"
        "at your organization.\n\n"
        "Use CO-STAR to build your prompt:\n\n"
        "  C  — What's the context/background?\n"
        "  O  — What's the objective of this post?\n"
        "  S  — What writing style fits LinkedIn?\n"
        "  T  — What tone do you want?\n"
        "  A  — Who is your target audience?\n"
        "  R  — What format should the output be?",
    0.6, 2.3, 5.6, 4.5,
    size=16, color=C_DARK)

rounded_rect(sl, 6.8, 1.65, 6.1, 5.5, C_BG2)
txb(sl, "Build Your CO-STAR Prompt", 7.0, 1.75, 5.7, 0.45,
    size=22, bold=True, color=C_PURPLE)

txb(sl, "C (Context):\n"
        "_________________________________\n\n"
        "O (Objective):\n"
        "_________________________________\n\n"
        "S (Style):\n"
        "_________________________________\n\n"
        "T (Tone):\n"
        "_________________________________\n\n"
        "A (Audience):\n"
        "_________________________________\n\n"
        "R (Response):\n"
        "_________________________________",
    7.0, 2.3, 5.7, 4.8,
    size=15, color=C_DGRAY)

txb(sl, "Time: 5 minutes  |  Compare your prompt and output with a neighbor!",
    0.4, 7.15, 12.5, 0.35,
    size=16, bold=True, color=C_ORANGE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION DIVIDER: GREEN PROMPTING
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(2, "Green Prompting",
                subtitle="How better prompts reduce environmental impact",
                accent=C_LEAF)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 12: ENVIRONMENTAL IMPACT OF AI
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "The Environmental Cost of AI",
           subtitle="Every query has an energy footprint", bar_color=C_LEAF)

# Left - the problem
rounded_rect(sl, 0.4, 1.65, 6.0, 5.5, RGBColor(0xFC, 0xE4, 0xEC))
txb(sl, "The Reality", 0.6, 1.75, 5.6, 0.45,
    size=22, bold=True, color=C_PINK)

txb(sl, "AI models run on massive data centers that\n"
        "consume significant energy and water.\n\n"
        "  A single ChatGPT query uses roughly\n"
        "  10x the energy of a Google search\n\n"
        "  Training GPT-4 used an estimated\n"
        "  50 GWh of electricity\n\n"
        "  Data centers account for ~1-1.5% of\n"
        "  global electricity consumption\n\n"
        "Every unnecessary re-prompt, vague query,\n"
        "or discarded output wastes energy.",
    0.6, 2.35, 5.6, 4.5,
    size=16, color=C_DARK)

# Right - the solution
rounded_rect(sl, 6.8, 1.65, 6.1, 5.5, RGBColor(0xE8, 0xF5, 0xE9))
txb(sl, "Better Prompts = Greener AI", 7.0, 1.75, 5.7, 0.45,
    size=22, bold=True, color=C_LEAF)

txb(sl, "When you prompt well, you:\n\n"
        "  Reduce the number of follow-up queries\n"
        "  — fewer round trips = less compute\n\n"
        "  Get usable output on the first try\n"
        "  — no wasted generation cycles\n\n"
        "  Use shorter, targeted prompts when simple\n"
        "  tasks don't need elaborate frameworks\n\n"
        "  Make intentional choices about when to\n"
        "  use AI vs. when you don't need it\n\n"
        "Efficient prompting is sustainable prompting.",
    7.0, 2.35, 5.7, 4.5,
    size=16, color=C_DARK)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 13: GREEN PROMPTING TIPS
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Green Prompting Best Practices",
           subtitle="Small changes, big environmental impact", bar_color=C_LEAF)

tips = [
    ("Be Specific the First Time",
     "Use a framework. Provide context. Vague prompts lead to multiple retries.",
     C_LEAF),
    ("Right-Size Your Prompt",
     "Not every task needs a 6-part framework. A simple question? Just ask it clearly.",
     C_TEAL),
    ("Batch Related Questions",
     "Combine related asks into one prompt instead of sending five separate messages.",
     C_BLUE),
    ("Evaluate Before Re-Prompting",
     "Can you edit the output yourself instead of regenerating? Often a small tweak is faster and greener.",
     C_PURPLE),
    ("Ask Yourself: Do I Need AI for This?",
     "Sometimes a quick search, a template, or your own expertise is the better tool.",
     C_ORANGE),
]

y_pos = 1.7
for title, desc, color in tips:
    rounded_rect(sl, 0.4, y_pos, 12.5, 1.0, C_BG2)
    rect(sl, 0.4, y_pos, 0.25, 1.0, color)
    txb(sl, title, 0.85, y_pos + 0.05, 11.8, 0.35,
        size=18, bold=True, color=color)
    txb(sl, desc, 0.85, y_pos + 0.42, 11.8, 0.55,
        size=15, color=C_DGRAY)
    y_pos += 1.12


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION DIVIDER: REFINING PROMPTS
# ═══════════════════════════════════════════════════════════════════════════════
section_divider(3, "Refining Prompts &\nEvaluating Outputs",
                subtitle="Iterating toward excellence",
                accent=C_BLUE)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 14: THE REFINEMENT CYCLE
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "The Prompt Refinement Cycle",
           subtitle="Great prompts rarely happen on the first try", bar_color=C_BLUE)

# Step cards in a flow
steps = [
    ("1", "Draft Your Prompt", "Use a framework to\nbuild your initial prompt", C_BLUE),
    ("2", "Run & Review", "Submit the prompt and\ncarefully read the output", C_PURPLE),
    ("3", "Identify Gaps", "What's missing? What's\nwrong? What's off-tone?", C_PINK),
    ("4", "Refine & Iterate", "Adjust the prompt based\non what you learned", C_ORANGE),
    ("5", "Validate", "Does the final output meet\nyour quality criteria?", C_GREEN),
]

x_start = 0.35
card_w = 2.35
gap = 0.2
for i, (num, title, desc, color) in enumerate(steps):
    x = x_start + i * (card_w + gap)
    rounded_rect(sl, x, 1.8, card_w, 2.6, C_BG2)
    rect(sl, x, 1.8, card_w, 0.55, color)
    txb(sl, f"Step {num}", x + 0.1, 1.83, card_w - 0.2, 0.45,
        size=18, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    txb(sl, title, x + 0.15, 2.45, card_w - 0.3, 0.4,
        size=16, bold=True, color=C_DARK, align=PP_ALIGN.CENTER)
    txb(sl, desc, x + 0.15, 2.9, card_w - 0.3, 1.3,
        size=14, color=C_DGRAY, align=PP_ALIGN.CENTER)

# Arrow hints between steps
for i in range(4):
    x_arrow = x_start + (i + 1) * (card_w + gap) - gap / 2 - 0.08
    txb(sl, ">", x_arrow, 2.7, 0.3, 0.4,
        size=28, bold=True, color=C_DGRAY, align=PP_ALIGN.CENTER)

# Refinement strategies
rounded_rect(sl, 0.4, 4.7, 12.5, 2.5, C_BG2)
txb(sl, "Refinement Strategies", 0.6, 4.8, 12.0, 0.4,
    size=20, bold=True, color=C_BLUE)

strategies = [
    ("Add more context", "If the output is too generic, provide specific details, examples, or constraints"),
    ("Adjust the format", "If the structure doesn't work, specify a different output format explicitly"),
    ("Change the role", "If the tone is wrong, try a different expert persona that better fits your needs"),
    ("Break it down", "If the output is too broad, split your prompt into smaller, focused sub-tasks"),
]
y_pos = 5.25
for title, desc in strategies:
    txb(sl, f"  {title}:", 0.6, y_pos, 3.0, 0.3,
        size=15, bold=True, color=C_BLUE)
    txb(sl, desc, 3.3, y_pos, 9.5, 0.3,
        size=15, color=C_DGRAY)
    y_pos += 0.38


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 15: EVALUATING OUTPUTS
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Evaluating AI Outputs",
           subtitle="Critical thinking is your superpower", bar_color=C_BLUE)

criteria = [
    ("Accuracy", "Is the information factually correct? Cross-check claims.",
     "Don't assume AI outputs are accurate — verify key facts", C_PINK),
    ("Relevance", "Does it actually address what you asked for?",
     "A well-written but off-topic answer is still a miss", C_PURPLE),
    ("Completeness", "Is anything important missing from the response?",
     "Check if all parts of your request were addressed", C_BLUE),
    ("Tone & Style", "Does it match the intended voice and audience?",
     "Read it as if your audience will — does it land right?", C_TEAL),
    ("Bias & Fairness", "Does the output show unintended bias or assumptions?",
     "Watch for stereotypes, one-sided perspectives, or exclusions", C_ORANGE),
]

y_pos = 1.7
for title, question, tip, color in criteria:
    rounded_rect(sl, 0.4, y_pos, 12.5, 1.0, C_BG2)
    rect(sl, 0.4, y_pos, 0.25, 1.0, color)
    txb(sl, title, 0.85, y_pos + 0.02, 2.5, 0.35,
        size=18, bold=True, color=color)
    txb(sl, question, 3.2, y_pos + 0.02, 5.5, 0.35,
        size=15, color=C_DARK)
    txb(sl, tip, 3.2, y_pos + 0.4, 9.5, 0.55,
        size=13, color=C_DGRAY, italic=True)
    y_pos += 1.12


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 16: INTERACTIVE ACTIVITY 3 — REFINE A PROMPT
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Interactive Activity: Refine & Evaluate",
           subtitle="Practice the refinement cycle with a partner",
           bar_color=C_ORANGE)

# Instructions panel
rounded_rect(sl, 0.4, 1.65, 6.0, 5.5, RGBColor(0xFD, 0xF0, 0xE7))
txb(sl, "The Exercise", 0.6, 1.75, 5.6, 0.45,
    size=22, bold=True, color=C_ORANGE)

txb(sl, "Step 1: Pick any framework you've learned\n"
        "today and write a prompt for this task:\n\n"
        "\"Create training materials for onboarding\n"
        " new employees at your organization.\"\n\n"
        "Step 2: Run your prompt in an AI tool\n\n"
        "Step 3: Evaluate the output using the\n"
        "5 criteria (accuracy, relevance,\n"
        "completeness, tone, bias)\n\n"
        "Step 4: Identify 2-3 specific improvements\n"
        "and refine your prompt\n\n"
        "Step 5: Run again and compare!",
    0.6, 2.3, 5.6, 4.8,
    size=16, color=C_DARK)

# Evaluation checklist
rounded_rect(sl, 6.8, 1.65, 6.1, 5.5, C_BG2)
txb(sl, "Evaluation Checklist", 7.0, 1.75, 5.7, 0.45,
    size=22, bold=True, color=C_PURPLE)

checklist = [
    "Is the output accurate and factual?",
    "Does it address all parts of my request?",
    "Is anything important missing?",
    "Does the tone fit the intended audience?",
    "Do I see any biased assumptions?",
    "What would I change in my prompt?",
    "What worked well that I should keep?",
    "Is this better than my first attempt?",
]

y_pos = 2.4
for item in checklist:
    txb(sl, f"   {item}", 7.0, y_pos, 5.7, 0.35,
        size=16, color=C_DARK)
    y_pos += 0.48

txb(sl, "Time: 8 minutes  |  Be ready to share what you changed and why!",
    0.4, 7.15, 12.5, 0.35,
    size=16, bold=True, color=C_ORANGE, italic=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 17: KEY TAKEAWAYS
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide()
header_bar(sl, "Key Takeaways", bar_color=C_PURPLE)

takeaways = [
    ("Frameworks give you structure",
     "RTF, CREATE, CO-STAR, CRISPE, RISEN, RACE — pick the one that fits your task",
     C_BLUE),
    ("Better prompts = better results, faster",
     "Specificity and context reduce ambiguity and rework",
     C_PURPLE),
    ("You stay in charge",
     "Good prompting keeps critical thinking and human judgment at the center",
     C_TEAL),
    ("Prompting can be greener",
     "Efficient prompts reduce compute waste and environmental impact",
     C_LEAF),
    ("Iterate and evaluate",
     "Use the refinement cycle and evaluation criteria to continuously improve",
     C_ORANGE),
]

y_pos = 1.7
for title, desc, color in takeaways:
    rounded_rect(sl, 0.4, y_pos, 12.5, 1.0, C_BG2)
    rect(sl, 0.4, y_pos, 0.25, 1.0, color)
    txb(sl, title, 0.85, y_pos + 0.05, 11.8, 0.35,
        size=20, bold=True, color=color)
    txb(sl, desc, 0.85, y_pos + 0.45, 11.8, 0.5,
        size=16, color=C_DGRAY)
    y_pos += 1.12


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 18: THANK YOU / Q&A
# ═══════════════════════════════════════════════════════════════════════════════
sl = add_slide(bg=C_PURPLE)
rect(sl, 0, 0, 0.55, 7.5, C_YELLOW)
txb(sl, "Thank You!", 0.85, 2.0, 12, 1.3,
    size=56, bold=True, color=C_WHITE)
txb(sl, "Questions & Discussion", 0.85, 3.4, 12, 0.8,
    size=32, color=C_LGRAY)
rect(sl, 0.85, 4.3, 4.5, 0.06, C_YELLOW)
txb(sl, "Keep practicing with frameworks — the more you use them,\n"
        "the more natural structured prompting becomes.",
    0.85, 4.7, 10, 1.0,
    size=20, color=C_LGRAY, italic=True)
txb(sl, "Remember: AI is a tool — your expertise drives the outcome.",
    0.85, 5.8, 10, 0.5,
    size=18, color=C_YELLOW, bold=True)


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
output_file = "Intermediate_AI_Prompting.pptx"
prs.save(output_file)
print(f"Presentation saved: {output_file}")
print(f"Total slides: {len(prs.slides)}")
