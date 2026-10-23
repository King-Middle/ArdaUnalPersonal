# -*- coding: utf-8 -*-
"""
Arda Ünal — Professional CV (Word .docx)
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()

# ── Page Setup ──
section = doc.sections[0]
section.top_margin = Cm(1.5)
section.bottom_margin = Cm(1.5)
section.left_margin = Cm(2.0)
section.right_margin = Cm(2.0)

# ── Style Defaults ──
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(10.5)
font.color.rgb = RGBColor(0x2D, 0x2D, 0x2D)
style.paragraph_format.space_before = Pt(0)
style.paragraph_format.space_after = Pt(2)
style.paragraph_format.line_spacing = 1.1

# ══════════════════════════════════════════════════════════════
#  COLOR PALETTE
# ══════════════════════════════════════════════════════════════
NAVY = RGBColor(0x1A, 0x23, 0x5B)       # section headings
DARK = RGBColor(0x22, 0x22, 0x22)        # body text
ACCENT = RGBColor(0x2E, 0x74, 0xB5)      # subtle accent
GRAY = RGBColor(0x66, 0x66, 0x66)        # dates / secondary text
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ══════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════

def set_cell_shading(cell, color_hex):
    """Set background color of a table cell."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def remove_paragraph_spacing(paragraph):
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)


def add_horizontal_line(doc, color='1A235B', thickness=1.5):
    """Add a colored horizontal line."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(6)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(int(thickness * 8)))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def add_section_heading(doc, text):
    """Add a styled section heading with bottom border."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text.upper())
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = NAVY
    r.font.name = 'Calibri'
    # Add bottom border
    add_horizontal_line(doc, '1A235B', 1)
    return p


def add_entry_header(doc, title, right_text=None):
    """Add a bold title with optional right-aligned date."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)
    
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'
    
    if right_text:
        r = p.add_run('\t' + right_text)
        r.font.size = Pt(10)
        r.font.color.rgb = GRAY
        r.font.name = 'Calibri'
        r.italic = True
        # Add right tab stop
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Cm(17.0), alignment=2)  # RIGHT aligned
    return p


def add_subtitle(doc, text):
    """Add italic subtitle line."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = ACCENT
    r.font.name = 'Calibri'
    return p


def add_bullet(doc, text):
    """Add a bullet point."""
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(1.0)
    # Clear auto-generated runs and add styled text
    p.clear()
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'
    return p


def add_body_text(doc, text):
    """Add a normal body paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    r.font.size = Pt(10)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'
    return p


# ══════════════════════════════════════════════════════════════
#  HEADER — NAME & CONTACT
# ══════════════════════════════════════════════════════════════

# Name
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(2)
r = p.add_run('ARDA ÜNAL')
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = NAVY
r.font.name = 'Calibri'

# Tagline
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(4)
r = p.add_run('Aspiring Computer Scientist & AI Researcher  |  IB Diploma Candidate')
r.font.size = Pt(11)
r.font.color.rgb = ACCENT
r.font.name = 'Calibri'

# Contact line
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(2)

contacts = [
    ('Istanbul, Turkey', None),
    ('+90 532 266 50 50', None),
    ('ardaunal1000@gmail.com', 'mailto:ardaunal1000@gmail.com'),
]

for i, (text, link) in enumerate(contacts):
    if i > 0:
        r = p.add_run('  |  ')
        r.font.size = Pt(9.5)
        r.font.color.rgb = GRAY
    r = p.add_run(text)
    r.font.size = Pt(9.5)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'

# Second contact line — LinkedIn & GitHub
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(0)
p.paragraph_format.space_after = Pt(2)

r = p.add_run('LinkedIn: ')
r.font.size = Pt(9.5)
r.font.color.rgb = GRAY
r.font.name = 'Calibri'
r = p.add_run('linkedin.com/in/arda-ünal-1ba497396')
r.font.size = Pt(9.5)
r.font.color.rgb = ACCENT
r.font.name = 'Calibri'

r = p.add_run('  |  ')
r.font.size = Pt(9.5)
r.font.color.rgb = GRAY

r = p.add_run('GitHub: ')
r.font.size = Pt(9.5)
r.font.color.rgb = GRAY
r.font.name = 'Calibri'
r = p.add_run('github.com/King-Middle')
r.font.size = Pt(9.5)
r.font.color.rgb = ACCENT
r.font.name = 'Calibri'

add_horizontal_line(doc, '1A235B', 2)

# ══════════════════════════════════════════════════════════════
#  ABOUT ME
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'About Me')

add_body_text(doc,
    "I'm an IB student at SEV American College in Istanbul with a deep curiosity for Computer Science "
    "and AI. I taught myself C# at fifteen and published a game on Steam, built a financial forecasting "
    "model using neural networks at Oxford, and co-developed a patented LSTM-based prediction system — "
    "all before graduating high school. What drives me isn't just building things that work; it's "
    "understanding why they sometimes don't. I care about making AI that's honest about its limitations, "
    "and I bring the same honesty to my own learning: I'd rather dig into a failure than take credit for "
    "a success I can't fully explain."
)

# ══════════════════════════════════════════════════════════════
#  EDUCATION
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Education')

add_entry_header(doc, 'SEV American College — IB Diploma Programme', '2021 – Present')
add_subtitle(doc, 'Istanbul, Turkey')
add_bullet(doc, 'IB Diploma Candidate with a focus on Mathematics, Computer Science, and Physics')
add_bullet(doc, 'Consistently pursuing independent projects alongside coursework to bridge theory with real-world application')

add_entry_header(doc, 'Immerse Education — AI & Machine Learning Programme', 'Summer 2024')
add_subtitle(doc, 'University of Oxford, United Kingdom')
add_bullet(doc, 'Spent the summer at Oxford learning to build neural networks and work with real data — not just textbook examples')
add_bullet(doc, 'Designed a financial forecasting bot that processed 5 years of stock data and 28 macroeconomic indicators')
add_bullet(doc, "Received a professor's reference letter after going beyond the curriculum to investigate why my model failed on real-world data")

add_entry_header(doc, 'Aziz Nesin Mathematics Village', 'Summer')
add_subtitle(doc, 'Şirince, Turkey')
add_bullet(doc, 'Lived and studied at a math camp where I combined Python programming with mathematical modelling')
add_bullet(doc, 'Discovered firsthand how mathematical reasoning and coding feed into each other')

# ══════════════════════════════════════════════════════════════
#  TECHNICAL SKILLS
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Technical Skills')

skills_data = [
    ('Programming Languages', 'Python, C#, Java'),
    ('AI / Machine Learning', 'Neural Networks, LSTM, Scikit-learn, TensorFlow/Keras, Data Preprocessing, Feature Engineering'),
    ('Game Development', 'Unity Engine, 3D Game Design, Physics Simulation'),
    ('Web Technologies', 'Web-based Application Development, REST APIs'),
    ('Data & Analytics', 'Pandas, NumPy, Financial Data Analysis, Time Series Forecasting'),
    ('Tools & Platforms', 'Git/GitHub, Steam (Publishing), Medium (Technical Writing)'),
    ('Concepts', 'OOP, Algorithms & Data Structures, Quantum Programming Fundamentals, Bias-Variance Tradeoff, Regularization'),
]

for label, value in skills_data:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f'{label}:  ')
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = NAVY
    r.font.name = 'Calibri'
    r = p.add_run(value)
    r.font.size = Pt(10)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'

# ══════════════════════════════════════════════════════════════
#  PROJECTS & EXPERIENCE
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Projects & Experience')

# Project 1: Flying Up
add_entry_header(doc, 'Flying Up — Indie Game Published on Steam', '2023')
add_subtitle(doc, 'Solo Developer | Unity / C#')
add_bullet(doc, 'Taught myself C# and Unity at 15 with no formal training — designed, built, and shipped a full 3D game on my own')
add_bullet(doc, 'Had to learn trigonometry and vector math months ahead of school because the physics engine wouldn\'t work without it')
add_bullet(doc, 'Went through the entire process: architecture, coding, playtesting, debugging, and finally hitting "Publish" on Steam')

# Project 2: Financial Forecasting Bot
add_entry_header(doc, 'AI Financial Forecasting Bot', '2024')
add_subtitle(doc, 'Immerse Education, Oxford | Python, ML/DL')
add_bullet(doc, 'Built a neural network that combined 5 years of stock data with 28 macroeconomic indicators (VIX, DXY, S&P 500, etc.)')
add_bullet(doc, 'The model trained beautifully but failed on real data — spent weeks understanding overfitting, distribution shift, and regularization')
add_bullet(doc, 'Wrote about what I learned from the failure on Medium, because the mistakes taught me more than the successes')

# Project 3: Patent System
add_entry_header(doc, 'LSTM-Based Stock Price Prediction System — Patent Application', '2024 – 2025')
add_subtitle(doc, 'Co-inventor | Python, LSTM, Web Platform')
add_bullet(doc, 'Co-built a multi-layered LSTM system that feeds 8 macroeconomic indicators alongside technical OHLCV data into predictions')
add_bullet(doc, 'Ran dozens of experiments to find the right combination of time windows, architecture depth, and hyperparameters')
add_bullet(doc, 'Delivered the system through a web platform so non-technical users could actually use it; patent application filed')

# Project 4: Quantum Programming
add_entry_header(doc, 'Quantum Programming Exploration', '2024')
add_bullet(doc, 'Dove into quantum computing out of curiosity — studied quantum state representation and its connections to linear algebra')
add_bullet(doc, 'Pushed my understanding of computation beyond the classical models I was used to')

# ══════════════════════════════════════════════════════════════
#  LEADERSHIP & ACTIVITIES
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Leadership & Activities')

add_entry_header(doc, 'President — Mathematics Competition Club', 'SEV American College')
add_bullet(doc, 'Took over a club of solo problem-solvers and turned it into a team — paired algebraic thinkers with geometric thinkers so they could learn from each other')
add_bullet(doc, "Our Purple Comet results improved, but what I'm prouder of is that students started debating math in the hallways between sessions")
add_bullet(doc, 'Learned that leadership in an intellectual space isn\'t about being the smartest person — it\'s about making everyone\'s thinking visible')

add_entry_header(doc, 'Member — Phoenix AI Club', 'SEV American College')
add_bullet(doc, 'Pushed for an ethical review step before we start any AI project — because building something cool isn\'t enough if you haven\'t asked who it affects')
add_bullet(doc, 'Led discussions on algorithmic bias and responsible AI, drawing from real cases, not just theory')

add_entry_header(doc, 'FRC Robotics Club — Community Outreach', 'SEV American College')
add_bullet(doc, 'Designed a robotics workshop for kindergarteners where they arranged physical command blocks to make robots move (one girl made hers "dance")')
add_bullet(doc, 'Realized that teaching a five-year-old what an algorithm is forces you to understand it better yourself')

add_entry_header(doc, 'School Magazine — Digital Citizenship Columnist')
add_bullet(doc, 'Wrote articles making the case that understanding technology isn\'t optional anymore — it\'s a civic responsibility')

add_entry_header(doc, 'Wharton Global High School Investment Competition')
add_bullet(doc, 'Competed in teams, building investment strategies from scratch — learned how financial analysis and strategic thinking intersect under pressure')

# ══════════════════════════════════════════════════════════════
#  INTERNSHIP
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Internship')

add_entry_header(doc, 'Professional Engineering Intern — Meyer Group', '2024')
add_bullet(doc, 'Got a firsthand look at how engineering actually works in industry — very different pace and priorities than personal projects')
add_bullet(doc, 'Worked alongside professional engineers and saw how academic concepts I\'d studied translate (and sometimes don\'t) into production environments')
add_bullet(doc, 'Came away with a much clearer picture of the kind of engineer I want to become')

# ══════════════════════════════════════════════════════════════
#  AWARDS & ACHIEVEMENTS
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Awards & Achievements')

awards = [
    ('University of Waterloo Mathematics Competition — School Champion (2 consecutive years)', 'Won the Cayley School Champion Medal from CEMC — both times after rebuilding my problem-solving methods from scratch'),
    ('Turkey Intelligence Foundation National Exam — Ranked 43rd Nationally', 'Placed 43rd out of thousands in a national analytical reasoning and math exam'),
    ('Royal Academy of Music — Grade 8 Certificate', 'Years of disciplined practice that taught me patience and persistence long before coding did'),
    ('Steam Game Publication', 'Published a complete game on Steam at 16 — solo, self-taught, from first line of code to final deployment'),
]

for title, desc in awards:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run('▸  ')
    r.font.color.rgb = ACCENT
    r.font.size = Pt(10)
    r = p.add_run(title)
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'
    
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(2)
    p2.paragraph_format.left_indent = Cm(0.7)
    r = p2.add_run(desc)
    r.font.size = Pt(9.5)
    r.font.color.rgb = GRAY
    r.font.name = 'Calibri'
    r.italic = True

# ══════════════════════════════════════════════════════════════
#  LANGUAGES
# ══════════════════════════════════════════════════════════════

add_section_heading(doc, 'Languages')

languages = [
    ('Turkish', 'Native speaker'),
    ('English', 'Fluent — educated in English (IB), studied at Oxford, competed internationally'),
]

for lang, level in languages:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f'{lang}:  ')
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = NAVY
    r.font.name = 'Calibri'
    r = p.add_run(level)
    r.font.size = Pt(10.5)
    r.font.color.rgb = DARK
    r.font.name = 'Calibri'

# ══════════════════════════════════════════════════════════════
#  SAVE
# ══════════════════════════════════════════════════════════════

output_dir = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(output_dir, 'Arda_Unal_CV.docx')
doc.save(output_path)
print(f'✅ CV saved to: {output_path}')
print(f'   File size: {os.path.getsize(output_path) / 1024:.1f} KB')
