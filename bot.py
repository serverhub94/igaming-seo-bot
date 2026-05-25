import os
import re
import io
import json
import asyncio
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application,
    MessageHandler,
    CommandHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

# ─── НАСТРОЙКИ ───────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
DEEPSEEK_KEY   = os.environ.get("DEEPSEEK_KEY", "")
ALLOWED_USER_IDS = []  # пример: [123456789]. Пусто = всем.

# ─── СОСТОЯНИЯ ДИАЛОГА ───────────────────────────────────────────────────────
(
    STATE_BRAND,
    STATE_GEO,
    STATE_LANG,
    STATE_PAGE_TYPE,
    STATE_WORDS,
    STATE_H2,
    STATE_FAQ,
    STATE_QTY,
    STATE_KEYWORDS,
    STATE_URLS,
    STATE_CONFIRM,
) = range(11)

# ─── GEO DATA ────────────────────────────────────────────────────────────────
GEO_DATA = {
    "DE": {"license": "GGL (Gemeinsame Glücksspielbehörde der Länder)",
           "rg": "BZgA (0800 137 27 00) · Spielerschutz.de · GluecksSpirale.de"},
    "IT": {"license": "ADM (Agenzia delle Dogane e dei Monopoli)",
           "rg": "GiocaResponsabile.it · SIPAC · Azzardopatia.it"},
    "LV": {"license": "IAUI / Curaçao GCB (OGL/2024/688/0234)",
           "rg": "Drosme.lv · Latvijas Atkarību profilakses centrs · 116123"},
    "AT": {"license": "Austrian Gaming Authority",
           "rg": "Spielerhilfe.at · Spielsuchthilfe.at"},
    "ES": {"license": "DGOJ", "rg": "FEJAR · Jugarbien.es"},
    "PL": {"license": "MF (Ministerstwo Finansów)",
           "rg": "Centrum Wsparcia dla Graczy · 19002"},
    "UA": {"license": "KRAIL", "rg": "Телефон довіри 0-800-213-800"},
    "CA": {"license": "Kahnawake Gaming Commission",
           "rg": "ConnexOntario · CAMH · BCLC GameSense"},
    "NZ": {"license": "Department of Internal Affairs NZ",
           "rg": "Problem Gambling Foundation NZ · 0800 654 655"},
    "FI": {"license": "Veikkaus / Finnish Gaming Authority",
           "rg": "Peluuri.fi · 0800 100 101"},
}

# ─── ПРОМТЫ ПО ТИПУ + ВЕРСИИ ─────────────────────────────────────────────────
# Формат: TYPE_VERSION → инструкции для системного промта
CONTENT_PROFILES = {

    # ── text_review ──────────────────────────────────────────────────────────
    "text_review_v2": {
        "label": "Гайд / Простой обзор",
        "words_default": "1500-2000",
        "h2_default": 6,
        "faq_default": 5,
        "instructions": """
CONTENT PROFILE: text_review V2 — Simple Guide / Short Review
Target: informational queries, top-of-funnel readers
Style: friendly, accessible, not technical
Structure:
  — Short intro (40-60 words, answer-first)
  — 6 H2 sections, each 150-250 words
  — 1 visual element per H2 (alternate table/list)
  — No heavy commercial blocks
  — 1 soft CTA
  — 5 FAQ questions
Tone: helpful, beginner-friendly, clear
DO NOT: complex tables, heavy bonus terms, deep technical details
""",
    },

    "text_review_v3": {
        "label": "Listicle / Top-N / Обзор казино",
        "words_default": "2500-3000",
        "h2_default": 8,
        "faq_default": 6,
        "instructions": """
CONTENT PROFILE: text_review V3 — Listicle / Top-N / Casino Review
Target: competitive SERP queries, comparison intent
Style: structured, scannable, comparison-heavy
Structure:
  — Strong answer-first intro (60-80 words)
  — 8 H2 sections, each 200-300 words
  — Heavy use of comparison tables and ranked lists
  — Top-N format where applicable (Top 5 features, Top 3 bonuses)
  — Commercial blocks: bonus, payments, withdrawal, mobile
  — Trust block: license, RNG, responsible gambling
  — 6 FAQ questions targeting People Also Ask
Tone: authoritative, comparative, commercially useful
MUST INCLUDE: comparison table vs competitors
""",
    },

    "text_review_v4": {
        "label": "Экспертное ревью / Флагманская статья",
        "words_default": "3500-4500",
        "h2_default": 11,
        "faq_default": 8,
        "instructions": """
CONTENT PROFILE: text_review V4 — Expert Review / Flagship Article
Target: high-intent, money keywords, featured snippet competition
Style: deep expert analysis, maximum E-E-A-T signals
Structure:
  — Comprehensive answer-first intro (80-100 words)
  — 11+ H2 sections, each 250-400 words
  — Every topic developed from 4 angles: What/Why/How/Action
  — Full commercial coverage: bonus+wagering, payments, withdrawal, KYC,
    mobile app, support, sportsbook vs casino, live betting, promo codes
  — Full trust coverage: license, RNG cert, responsible gambling,
    real reviews, complaints, security, SSL, payout transparency
  — Author bio with credentials
  — 8 FAQ questions covering all PAA patterns
  — Internal linking suggestions
  — Pros/cons table mandatory
Tone: journalistic expert, maximum credibility, zero hype
MUST INCLUDE: pros/cons table, author bio, comparison table, all commercial blocks
""",
    },

    # ── mono ─────────────────────────────────────────────────────────────────
    "mono_v2": {
        "label": "Простой монобренд / Слот",
        "words_default": "2000-2500",
        "h2_default": 7,
        "faq_default": 5,
        "instructions": """
CONTENT PROFILE: mono V2 — Simple Monobrand / Slot Review
Target: branded queries, navigational + informational intent
Style: focused on one brand/game, clear and direct
Structure:
  — Brand-focused intro (60-80 words)
  — 7 H2 sections dedicated to single brand/game:
    • What is [Brand] / How it works
    • Key features / Game mechanics
    • Bonuses and promotions
    • Payment methods
    • Mobile experience
    • License and safety
    • FAQ
  — 1 visual element per H2
  — 5 FAQ questions about the brand
Tone: informative, brand-focused, trustworthy
DO NOT: compare with many competitors (1 brief comparison table max)
""",
    },

    "mono_v3": {
        "label": "Детальный монобренд с анализом конкурентов",
        "words_default": "3000-3500",
        "h2_default": 10,
        "faq_default": 7,
        "instructions": """
CONTENT PROFILE: mono V3 — Detailed Monobrand + Competitor Analysis
Target: branded + semi-branded queries, commercial intent
Style: deep brand analysis with competitive positioning
Structure:
  — Strong branded intro with key USPs (80 words)
  — 10 H2 sections:
    • Brand overview + quick facts table
    • Game/product mechanics in depth
    • Bonus structure + wagering analysis
    • Payment methods + withdrawal speed table
    • Mobile app / browser experience
    • Competitor comparison table (brand vs 3 competitors)
    • License, RNG, security
    • Real user reviews + pros/cons
    • Responsible gambling
    • FAQ (7 questions)
  — Full entity coverage from competitor URL analysis
  — Content gap exploitation from SERP analysis
Tone: expert, brand-authoritative, commercially complete
MUST INCLUDE: competitor comparison table, pros/cons, full commercial blocks
""",
    },

    # ── fallback ─────────────────────────────────────────────────────────────
    "default": {
        "label": "Casino Review (стандарт)",
        "words_default": "3000-3500",
        "h2_default": 8,
        "faq_default": 7,
        "instructions": "",
    },
}

def get_profile(tz_type: str, page_type: str) -> dict:
    """Match ТЗ type+version to content profile."""
    if not tz_type:
        return CONTENT_PROFILES["default"]
    
    t = tz_type.lower().strip()
    
    # Нормализуем ключ
    mapping = {
        "text_review_v2": "text_review_v2",
        "text_review v2": "text_review_v2",
        "text_review v3": "text_review_v3",
        "text_review_v3": "text_review_v3",
        "text_review v4": "text_review_v4",
        "text_review_v4": "text_review_v4",
        "mono_v2": "mono_v2",
        "mono v2": "mono_v2",
        "text_mono_v2": "mono_v2",
        "text_mono v2": "mono_v2",
        "mono_v3": "mono_v3",
        "mono v3": "mono_v3",
        "text_mono_v3": "mono_v3",
        "text_mono v3": "mono_v3",
        "text_mono_v4": "text_review_v4",
        "text_mono v4": "text_review_v4",
    }
    
    for pattern, profile_key in mapping.items():
        if pattern in t:
            return CONTENT_PROFILES[profile_key]
    
    # Попробуем по компонентам
    if "v4" in t:
        return CONTENT_PROFILES["text_review_v4"]
    if "v3" in t and "mono" in t:
        return CONTENT_PROFILES["mono_v3"]
    if "v3" in t:
        return CONTENT_PROFILES["text_review_v3"]
    if "v2" in t and "mono" in t:
        return CONTENT_PROFILES["mono_v2"]
    if "v2" in t:
        return CONTENT_PROFILES["text_review_v2"]
    
    return CONTENT_PROFILES["default"]



PAGE_TYPES = [
    "Casino Review", "Game Review", "Bonus Page",
    "Sportsbook Review", "How to Play Guide", "Comparison Page",
    "Landing Page", "Brand Page", "FAQ Page",
]

# ─── МАСТЕР-ПРОМТ v3 ─────────────────────────────────────────────────────────
SYSTEM_PROMPT = """
════════════════════════════════════════════════════
iGAMING SEO CONTENT GENERATION SYSTEM — v4 FULL
════════════════════════════════════════════════════

ROLE
You are a senior iGaming SEO content strategist (enterprise level).
Create content that:
  — Maximally satisfies search intent
  — Covers ALL semantic entities Google considers mandatory
  — Complies with Helpful Content / E-E-A-T
  — Reads as human-written expert content
  — Ranks through semantic completeness, NOT keyword stuffing

━━━ CORE PRINCIPLE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DO NOT optimize for keyword density.
DO NOT think "how many times does this keyword appear".
THINK: "what does Google consider mandatory to fully cover this intent?"

Optimize for:
  — Intent satisfaction (does the page answer EVERY user question?)
  — Semantic entity coverage (all expected concepts present?)
  — Topical completeness (no missing subtopics?)
  — Entity relevance (right entities for this brand/geo/type?)
  — Commercial usefulness (practical info for real decisions?)
  — Trust signals (regulation, verification, transparency?)
  — UX scanability (easy to read, scan, navigate?)

Exact-match keywords = LAST step, not first.
Primary goal: cover intent BETTER than every SERP competitor.

━━━ SERP ANALYSIS — FULL PROTOCOL ━━━━━━━━━━━━━━━━━
STEP 1 — If competitor URLs provided, analyze each:
  EXTRACT:
  — Title, H1, Meta Description
  — Full H2/H3/H4 heading structure
  — Word count estimate
  — Semantic entities present (what concepts are covered?)
  — Commercial blocks: bonus, payments, withdrawal, promo codes,
    app, registration, sportsbook vs casino, live betting, support
  — Trust blocks: license number, RNG cert, responsible gambling,
    KYC process, real reviews, complaints, security, SSL, payouts
  — FAQ questions and answers
  — Featured snippet / answer block candidates
  — People Also Ask patterns
  — Internal linking patterns
  — Content gaps (what's missing vs other competitors?)

STEP 2 — Build Competitor Entity Map:
  — Which semantic entities appear in ALL top pages? (mandatory)
  — Which appear in 50%+ pages? (important)
  — Which appear in only 1-2 pages? (differentiation opportunity)
  — What subtopics does NO competitor cover? (content gap)

STEP 3 — Calculate SERP Median (NOT hard constants):
  — Median word count for this page type
  — Median H2 count / H3 count
  — Median FAQ question count
  — Median intro length (words)
  — Median paragraph depth (sentences per paragraph)
  — Median sentence length
  — Unique vocabulary ratio
  — Trust section presence rate (%)
  — Commercial block density
  — Internal linking patterns
  USE ONLY pages of same type (review/bonus/landing etc.)
  DO NOT mix different intent types in median.

If no URLs provided:
  → Perform live SERP analysis for the primary keyword in target geo
  → Apply same extraction protocol to top-10 organic results

━━━ MANDATORY ENTITY COVERAGE ━━━━━━━━━━━━━━━━━━━━━
DO NOT think "mention this keyword N times".
THINK: "these specific entities MUST be explained and covered":

CORE entities (always mandatory):
  — Brand/game name, developer/operator, founding year
  — Core product mechanics (how it works in practice)
  — Key differentiators vs competitors

COMMERCIAL entities (cover ALL that apply):
  — Welcome bonus + wagering requirements (exact numbers)
  — Deposit methods (list with minimums)
  — Withdrawal speed (by method, with actual timeframes)
  — Promo codes / no-deposit bonus / free spins
  — Mobile app (iOS + Android, or browser-based)
  — Registration process (step-by-step)
  — Sportsbook vs casino difference (if both exist)
  — Live betting / live casino features
  — Customer support (channels, languages, hours)
  — Verification / KYC process

TRUST entities (ALL mandatory, never skip):
  — Exact license name + number + issuing authority
  — RNG certification (which lab, standard)
  — Responsible gambling resources (local, geo-specific)
  — Real user reviews sentiment
  — Complaints handling process
  — Security measures (SSL, 2FA, encryption)
  — Payout transparency (published RTP ranges)
  — Account verification requirements

TOPICAL subtopics (extract from SERP, cover top-recurring ones):
  — FAQ patterns from People Also Ask
  — Featured snippet answer blocks (40-60 word definitions)
  — Comparison subtopics (vs main competitors)
  — Legality in specific geo (local laws, license requirements)

━━━ QUALITY RESTRICTIONS (STRICT) ━━━━━━━━━━━━━━━━━
Apply ALL of these — no exceptions:

KEYWORD RULES:
  — Keyword density <= 2.5-3% total
  — Exact-match: H1 (required) + intro + 1-2 strategic H2 only
  — NO forced exact-match insertions anywhere else
  — Use semantic variants, synonyms, related phrases throughout
  — Primary keyword in first 20-30% of text (naturally)

READABILITY RULES:
  — Long sentences (25+ words): < 5% of all sentences
  — Paragraph scanability: 2-5 sentences per paragraph max
  — Clear heading hierarchy (H1 > H2 > H3, never skip levels)
  — Answer-first structure in first 20-30% of text
  — No шаблонный AI-style content ("In the world of...")
  — High semantic consistency (related terms cluster naturally)
  — NO: guaranteed wins, misleading claims, hype language
  — NO: robotic phrasing, empty filler, repetitive structures

NOTE for RU/UA/Slavic languages:
  DO NOT use Flesch readability as primary KPI — it is noisy.
  Priority order: readability → structure → semantic completeness → UX scanning.

━━━ INTENT-FIRST CONTENT GENERATION ━━━━━━━━━━━━━━━
Generate through INTENT + ENTITY COVERAGE, not keyword counts.

WRONG approach: "mention 'casino bonus' 7 times"
RIGHT approach: "cover these mandatory entities about bonuses:
  • exact welcome bonus amount and structure
  • wagering requirements (number + explanation)
  • which games count toward wagering
  • time limit for bonus completion
  • max bet with active bonus
  • how to activate (promo code or automatic)
  • free spins details (if applicable)"

For every major topic, ask:
  1. What does the user ACTUALLY want to know?
  2. What would an expert explain that competitors miss?
  3. What trust signal proves this information is reliable?
  4. What practical action should the reader take?

━━━ FEATURED SNIPPET OPTIMIZATION ━━━━━━━━━━━━━━━━━
In first 20-30% of text, add ONE block targeting:
  — Featured Snippets (concise definition 40-60 words)
  — AI Overviews (direct factual answer)
  — People Also Ask (question + short answer format)
  — Passage Ranking (self-contained explanatory paragraph)

Rotate formats across sections:
  Definition block → Step-by-step list → Comparison table → Q&A

━━━ HEADING DEPTH RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
Each H2 section: minimum 200-350 words.
Develop every topic from FOUR mandatory angles:
  1. WHAT — define and contextualize
  2. WHY — why does this matter to THIS reader right now?
  3. HOW — practical mechanics, real numbers, step-by-step
  4. ACTION — what should reader do with this information?

H3 subsections: specific detail layers, NOT repetition of H2.
NO thin sections: every H2 must earn its place with real content.

━━━ SECTION INTRO + VISUAL ELEMENT RULES ━━━━━━━━━━
Every H2 MUST follow this 3-step structure:

  STEP 1 — INTRO (1-2 sentences, MANDATORY)
    — Frame the topic + why it matters NOW to this reader
    — NEVER start H2 directly with table or list
    — NEVER use "In this section we will..."

  STEP 2 — VISUAL ELEMENT (MANDATORY, ONE per H2)
    TABLE    → comparisons, specs, rates, limits, payment methods
    BULLETS  → features, pros/cons, requirements (use: - item)
    NUMBERED → processes, steps, rankings (use: 1. step)
    Min 4 items per list · Min 3 data rows per table (+header)

  STEP 3 — BODY TEXT (after visual, MANDATORY)
    — Expand with context, nuance, practical advice
    — Add NEW information not shown in the visual
    — Minimum 2 paragraphs

  FORMAT ROTATION (mandatory):
    Never same visual format in two consecutive H2 sections.
    Alternate: table → bullets → numbered → table → ...

━━━ E-E-A-T REQUIREMENTS ━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Expert tone throughout (no hedging, no filler)
  — Trust signal in EVERY commercial section
  — Exact license reference (name + number + authority)
  — Realistic expectations (no hype, no guaranteed wins)
  — Practical user guidance (what to actually do)
  — Transparent pros AND cons (not just positives)
  — Responsible gambling with geo-specific resources
  — Author expertise block (name + credentials + experience)
  — Update freshness signal (last updated date)
  — External authority references where applicable

━━━ INTERNAL LINKING LOGIC ━━━━━━━━━━━━━━━━━━━━━━━━
  — Contextual links only where naturally relevant
  — Commercial funnel: bonus → registration → deposit → withdrawal
  — Entity links: game → casino → bonus → payments
  — DO NOT insert links artificially
  — Anchor text: descriptive, not over-optimized

━━━ POST-GENERATION QUALITY CHECK ━━━━━━━━━━━━━━━━━
Before outputting, mentally verify ALL:

  INTENT:
  ✓ Does the text directly answer the primary search query?
  ✓ Would a real user find everything they need here?
  ✓ Is the answer-first structure present in first 20-30%?

  ENTITIES:
  ✓ All commercial entities covered? (bonus/payments/withdrawal/app/support)
  ✓ All trust entities covered? (license/RNG/RG/KYC/security)
  ✓ All topical subtopics from SERP present?
  ✓ No missing FAQ patterns from People Also Ask?

  QUALITY:
  ✓ Zero keyword stuffing or unnatural insertions?
  ✓ No AI-style filler or robotic phrasing?
  ✓ Long sentences < 5% of total?
  ✓ Every H2 has intro + visual element + body text?
  ✓ Visual formats alternate across sections?
  ✓ Every H2 has 200-350+ words of real content?
  ✓ Pros/cons table present (for review pages)?
  ✓ Author bio + date at end?
  ✓ Responsible gambling with local resources?

  IF ANYTHING FAILS:
  DO NOT force-insert a keyword.
  DO add a meaningful semantic block covering the missing entity.

━━━ OUTPUT FORMAT (STRICT — used for .docx conversion) ━━━
# H1 Title (exact-match keyword naturally included)

Intro paragraph (60-80 words, answer-first, featured snippet candidate).

## H2 Section Title

Intro sentence(s) framing why this matters.

| Column 1 | Column 2 | Column 3 |
|---|---|---|
| Row 1 | Data | Data |
| Row 2 | Data | Data |

Body text paragraphs expanding on the table.

- Bullet item (min 4 items)
- Bullet item
- Bullet item
- Bullet item

1. Step one (min 4 steps)
2. Step two
3. Step three
4. Step four

**FAQ question in bold?**
Direct answer in 30-60 words.

---
*Author: [Full Name] — [specific credentials]. Last updated: [Month Year].*
*Responsible gambling: [geo-specific resources]. 18+*

═══ MULTI-VERSION SEPARATOR ═══
Use: === VERSION N === between versions.
Each version: unique H1 · different H2 structure · different intro angle.
Max 20% phrasing overlap between any two versions.
"""


# ─── MARKDOWN → DOCX ─────────────────────────────────────────────────────────
def markdown_to_docx(markdown_text: str) -> bytes:
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    def add_shading(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)

    def parse_inline(paragraph, text):
        parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            elif part.startswith('*') and part.endswith('*'):
                run = paragraph.add_run(part[1:-1])
                run.italic = True
            else:
                if part:
                    paragraph.add_run(part)

    def add_table(rows_data):
        data_rows = [r for r in rows_data if not re.match(r'^\|[-| :]+\|$', r.strip())]
        if len(data_rows) < 2:
            return
        def parse_row(row_str):
            cells = row_str.strip().strip('|').split('|')
            return [c.strip() for c in cells]
        parsed = [parse_row(r) for r in data_rows]
        col_count = max(len(r) for r in parsed)
        table = doc.add_table(rows=len(parsed), cols=col_count)
        table.style = 'Table Grid'
        for i, row_data in enumerate(parsed):
            row = table.rows[i]
            for j in range(col_count):
                cell = row.cells[j]
                text = row_data[j] if j < len(row_data) else ''
                cell.text = text
                if i == 0:
                    for run in cell.paragraphs[0].runs:
                        run.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_shading(cell, '2E75B6')
                elif i % 2 == 0:
                    add_shading(cell, 'EBF3FA')
        doc.add_paragraph()

    lines = markdown_text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i]

        # Разделитель версий
        if re.match(r'^=+\s*VERSION\s*\d+\s*=+$', line.strip(), re.I):
            p = doc.add_paragraph()
            run = p.add_run(f'\n{"═" * 50}\n{line.strip()}\n{"═" * 50}')
            run.bold = True
            run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            doc.add_paragraph()
            i += 1
            continue

        if re.match(r'^---+$', line.strip()):
            p = doc.add_paragraph('─' * 60)
            p.paragraph_format.space_after = Pt(4)
            i += 1
            continue

        if line.startswith('# ') and not line.startswith('## '):
            p = doc.add_heading(line[2:].strip(), level=1)
            run = p.runs[0] if p.runs else p.add_run(line[2:].strip())
            run.font.size = Pt(22)
            run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
            i += 1; continue

        if line.startswith('## ') and not line.startswith('### '):
            p = doc.add_heading(line[3:].strip(), level=2)
            run = p.runs[0] if p.runs else p.add_run(line[3:].strip())
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            i += 1; continue

        if line.startswith('### '):
            p = doc.add_heading(line[4:].strip(), level=3)
            run = p.runs[0] if p.runs else p.add_run(line[4:].strip())
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            i += 1; continue

        if line.startswith('|'):
            table_rows = []
            while i < len(lines) and lines[i].startswith('|'):
                table_rows.append(lines[i])
                i += 1
            add_table(table_rows)
            continue

        if re.match(r'^[-*•] ', line):
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
            parse_inline(p, re.sub(r'^[-*•] ', '', line).strip())
            i += 1; continue

        if re.match(r'^\d+\. ', line):
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_after = Pt(2)
            parse_inline(p, re.sub(r'^\d+\. ', '', line).strip())
            i += 1; continue

        if not line.strip():
            i += 1; continue

        if re.match(r'^\*\*.*\*\*$', line.strip()):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line.strip().strip('*'))
            run.bold = True
            i += 1; continue

        if re.match(r'^\*.*\*$', line.strip()):
            p = doc.add_paragraph()
            run = p.add_run(line.strip().strip('*'))
            run.italic = True
            run.font.size = Pt(10)
            i += 1; continue

        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        parse_inline(p, line.strip())
        i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── ГЕНЕРАЦИЯ ТЕКСТА ────────────────────────────────────────────────────────
def generate_content(params: dict) -> str:
    client = OpenAI(api_key=DEEPSEEK_KEY, base_url="https://api.deepseek.com")

    geo_code = params.get("geo", "??")
    geo_info = GEO_DATA.get(geo_code, {})

    # Строим промт из параметров
    urls_block = ""
    if params.get("urls"):
        url_list = "\n".join(f"  - {u}" for u in params["urls"])
        urls_block = f"""
COMPETITOR URLs TO ANALYZE:
{url_list}
For each URL extract: Title, H1, H2/H3 structure, semantic entities,
commercial blocks, trust signals, FAQ patterns, content gaps.
Build unified entity map and identify differentiation opportunities.
"""
    else:
        urls_block = f'NO URLS PROVIDED → Perform live SERP analysis for "{params.get("keyword", params.get("brand"))}" in {params.get("geo_name", geo_code)}.'

    keywords_block = ""
    if params.get("keywords"):
        kw_list = "\n".join(f"  - {k}" for k in params["keywords"])
        keywords_block = f"""
EXACT KEYWORDS TO USE:
Primary keyword (in H1 + intro + 1-2 H2): {params["keywords"][0]}
Secondary keywords (weave naturally, no stuffing):
{kw_list[kw_list.find(chr(10))+1:] if len(params["keywords"]) > 1 else "  (none — use semantic variants)"}
"""

    qty = params.get("qty", 1)
    multi_block = ""
    if qty > 1:
        multi_block = f"""
GENERATE {qty} UNIQUE VERSIONS:
  — Different H1, H2 structure, intro angle per version
  — Same mandatory entities, different perspectives
  — Max 20% phrasing overlap between versions
  — Separate versions with: === VERSION N ===
"""

    # Определяем профиль контента по типу ТЗ
    tz_type = params.get("tz_type", "") or ""
    profile = get_profile(tz_type, params.get("page_type", ""))
    profile_instructions = profile.get("instructions", "")

    # Применяем дефолты профиля если не заданы явно
    words = params.get("words") or profile["words_default"]
    h2    = params.get("h2")    or profile["h2_default"]
    faq   = params.get("faq")   or profile["faq_default"]

    user_prompt = f"""
CONTENT TASK:
Brand / Game:  {params.get("brand", "—")}
Geo:           {params.get("geo_name", geo_code)} ({geo_code})
Language:      {params.get("lang", "English")}
Page Type:     {params.get("page_type", "Casino Review")}
Content Type:  {tz_type.upper() if tz_type else "Standard"} — {profile["label"]}
Target words:  {words} per version
H2 sections:   {h2}
FAQ questions: {faq}
License:       {geo_info.get("license", "relevant regulator")}
Resp. Gambling:{geo_info.get("rg", "local resources")}

{profile_instructions}
{keywords_block}
{urls_block}
{multi_block}

Generate the complete article(s) in {params.get("lang", "English")}.
Follow all system rules: section intro → visual element → body text.
Output directly — no meta-commentary.
"""

    response = client.chat.completions.create(
        model="deepseek-chat",
        max_tokens=8000,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_prompt}
        ]
    )
    return response.choices[0].message.content


# ─── КОМАНДЫ ─────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "👋 *iGaming SEO Content Bot v5*\n\n"
        "📋 Просто вставь ТЗ в формате:\n"
        "`Количество текстов: 5`\n"
        "`Тип генерации: V4`\n"
        "`Тип: text_mono`\n"
        "`ГЕО + Язык: SK, Словацкий`\n\n"
        "`Конкуренты`\n"
        "`https://site1.com`\n\n"
        "`Ключи:`\n"
        "`main keyword`\n\n"
        "Или используй команды:\n"
        "• /new — пошаговый мастер\n"
        "• /types — все типы контента\n\n"
        "Результат — *.docx* файл 📎",
        parse_mode="Markdown"
    )

async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Отменено. Начни заново: /new или /quick",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


async def cmd_types(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show all content types and versions."""
    await update.message.reply_text(
        "📋 *Типы контента — шпаргалка:*\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "*text_review + V2*\n"
        "→ Гайды, простые обзоры\n"
        "→ 1500-2000 слов, 6 H2, 5 FAQ\n\n"
        "*text_review + V3*\n"
        "→ Listicle Top-N, обзоры казино\n"
        "→ 2500-3000 слов, 8 H2, 6 FAQ\n\n"
        "*text_review + V4*\n"
        "→ Экспертные ревью, флагманы\n"
        "→ 3500-4500 слов, 11 H2, 8 FAQ\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "*mono + V2* (text_mono)\n"
        "→ Простой монобренд / слот\n"
        "→ 2000-2500 слов, 7 H2, 5 FAQ\n\n"
        "*mono + V3* (text_mono)\n"
        "→ Детальный монобренд + конкуренты\n"
        "→ 3000-3500 слов, 10 H2, 7 FAQ\n\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "*Как указать в ТЗ:*\n"
        "`Тип генерации: V4`\n"
        "`Тип: text_mono`\n\n"
        "Бот сам подберёт объём, H2 и FAQ.",
        parse_mode="Markdown"
    )


# ─── БЫСТРЫЙ РЕЖИМ (/quick) ──────────────────────────────────────────────────
async def cmd_quick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "⚡ *Быстрый режим*\n\n"
        "Напиши одной строкой:\n"
        "`Бренд, Гео, Язык, Тип, Слова, Кол-во текстов`\n\n"
        "Пример:\n"
        "`GGBet, LV, латышский, Casino Review, 3000, 2`",
        parse_mode="Markdown"
    )

def parse_tz_format(text: str):
    """Parse ТЗ block format from image/message. Returns params dict or None."""
    lines = [l.strip() for l in text.strip().splitlines()]
    is_tz = any(
        "количество текстов" in l.lower() or
        "тз на генерацию" in l.lower() or
        "тип генерации" in l.lower()
        for l in lines
    )
    if not is_tz:
        return None

    params = {
        "brand": None, "geo": "??", "geo_name": "Unknown",
        "lang": "English", "page_type": "Casino Review",
        "words": "3000-3500", "h2": 8, "faq": 7,
        "qty": 1, "keywords": [], "urls": [],
        "tz_type": None, "gen_version": None,
    }

    PAGE_MAP = {
        "text_mono": "Casino Review", "mono": "Casino Review",
        "casino review": "Casino Review", "game review": "Game Review",
        "bonus": "Bonus Page", "landing": "Landing Page",
        "sportsbook": "Sportsbook Review", "how to play": "How to Play Guide",
        "comparison": "Comparison Page", "brand": "Brand Page",
    }

    mode = None
    for line in lines:
        ll = line.lower().strip()
        if not ll:
            continue

        # Количество текстов
        m = re.search(r'количество текстов[:\s]+(\d+)', ll)
        if m:
            params["qty"] = int(m.group(1))
            continue

        # Тип генерации: V4, V3 и т.д.
        m = re.search(r'тип генерации[:\s]+(.+)', ll)
        if m:
            params["gen_version"] = m.group(1).strip().lower()
            # Сохраняем для комбинирования с типом
            continue

        # Тип: text_mono, text_review и т.д. — объединяем с версией
        m = re.search(r'^тип[:\s]+(.+)', ll)
        if m:
            raw_type = m.group(1).strip().lower()
            gen_ver = params.get("gen_version", "")
            if gen_ver:
                params["tz_type"] = f"{raw_type}_{gen_ver}"
            else:
                params["tz_type"] = raw_type
            # Маппинг на page_type
            page_map = {
                "text_mono": "Casino Review", "mono": "Casino Review",
                "text_review": "Casino Review", "casino review": "Casino Review",
                "game review": "Game Review", "bonus": "Bonus Page",
                "landing": "Landing Page", "sportsbook": "Sportsbook Review",
            }
            for k, v in page_map.items():
                if k in raw_type:
                    params["page_type"] = v
                    break
            continue

        # Тип страницы
        m = re.search(r'^тип[:\s]+(.+)', ll)
        if m:
            raw = m.group(1).strip()
            for k, v in PAGE_MAP.items():
                if k in raw:
                    params["page_type"] = v
                    break
            else:
                params["page_type"] = raw.title()
            continue

        # ГЕО + Язык: SK, Словацкий  или  ГЕО: SK  или  Язык: Словацкий
        m = re.search(r'гео.*?язык[:\s]+([a-zA-Z]{2})[,\s]+(.+)', ll)
        if m:
            params["geo"] = m.group(1).upper()
            params["geo_name"] = m.group(1).upper()
            params["lang"] = line.split(',')[-1].strip().title() if ',' in line else m.group(2).strip().title()
            continue
        m = re.search(r'гео[:\s]+([a-zA-Z]{2})', ll)
        if m:
            params["geo"] = m.group(1).upper()
            params["geo_name"] = m.group(1).upper()
            continue
        m = re.search(r'язык[:\s]+(.+)', ll)
        if m:
            params["lang"] = m.group(1).strip().title()
            continue

        # Объём слов
        m = re.search(r'(слов|words)[:\s]+(\d[\d\-]+)', ll)
        if m:
            params["words"] = m.group(2)
            continue

        # Бренд явно
        m = re.search(r'^(бренд|игра|brand)[:\s]+(.+)', ll)
        if m:
            params["brand"] = m.group(2).strip()
            continue

        # H2
        m = re.search(r'h2[:\s]+(\d+)', ll)
        if m:
            params["h2"] = int(m.group(1))
            continue

        # FAQ
        m = re.search(r'faq[:\s]+(\d+)', ll)
        if m:
            params["faq"] = int(m.group(1))
            continue

        # Переключение режимов
        if re.search(r'конкурент', ll):
            mode = "urls"
            continue
        if re.match(r'ключи[:\s]?$', ll) or ll == "keywords:":
            mode = "keywords"
            continue

        # URL конкурентов
        if line.startswith("http"):
            params["urls"].append(line.strip())
            if not params["brand"] and params["urls"]:
                m2 = re.search(r'https?://(?:www\.)?([^/\.]+)', line)
                if m2:
                    params["brand"] = m2.group(1).title()
            continue

        # Ключевые слова
        if mode == "keywords" and line and not line.startswith("http"):
            skip = any(x in ll for x in [
                "конкурент", "ключи", "тип", "гео", "количество",
                "язык", "слов", "h2", "faq", "бренд"
            ])
            if not skip:
                params["keywords"].append(line.strip())
            continue

    # Бренд из ключей если не найден
    if not params["brand"]:
        if params["keywords"]:
            params["brand"] = params["keywords"][0].title()
        elif params["urls"]:
            m3 = re.search(r'https?://(?:www\.)?([^/\.]+)', params["urls"][0])
            params["brand"] = m3.group(1).title() if m3 else "Casino"
        else:
            params["brand"] = "Casino"

    return params


async def handle_quick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()

    # Пробуем сначала распарсить как ТЗ-формат
    tz_params = parse_tz_format(text)
    if tz_params:
        # Показываем что распарсили
        kw_preview = "\n".join(f"  • {k}" for k in tz_params["keywords"][:5])
        if len(tz_params["keywords"]) > 5:
            kw_preview += f"\n  ...и ещё {len(tz_params['keywords'])-5}"
        url_preview = "\n".join(f"  • {u}" for u in tz_params["urls"])

        # Определяем профиль для отображения
        tz_profile = get_profile(tz_params.get("tz_type",""), tz_params.get("page_type",""))
        profile_label = tz_profile["label"]
        tz_type_display = tz_params.get("tz_type","standard").upper()

        await update.message.reply_text(
            f"📋 *ТЗ распознано:*\n\n"
            f"🏷 Бренд: `{tz_params['brand']}`\n"
            f"📍 Гео: `{tz_params['geo']}` · Язык: `{tz_params['lang']}`\n"
            f"📄 Тип: `{tz_type_display}` — {profile_label}\n"
            f"📋 Версий: `{tz_params['qty']}`\n"
            f"📝 Слов: `{tz_profile['words_default']}`\n"
            f"🔢 H2: `{tz_profile['h2_default']}` · FAQ: `{tz_profile['faq_default']}`\n\n"
            f"🔑 Ключей: {len(tz_params['keywords'])}\n{kw_preview}\n\n"
            f"🔗 URL конкурентов: {len(tz_params['urls'])}\n{url_preview}\n\n"
            f"⏳ Запускаю генерацию...",
            parse_mode="Markdown"
        )
        await run_generation(update, context, tz_params)
        return

    # Иначе — быстрый формат через запятую
    parts = [p.strip() for p in text.split(',')]
    params = {
        "brand":     parts[0] if len(parts) > 0 else "Brand",
        "geo":       parts[1].upper() if len(parts) > 1 else "??",
        "geo_name":  parts[1] if len(parts) > 1 else "Unknown",
        "lang":      parts[2] if len(parts) > 2 else "English",
        "page_type": parts[3] if len(parts) > 3 else "Casino Review",
        "words":     parts[4] if len(parts) > 4 else "3000-3500",
        "qty":       int(parts[5]) if len(parts) > 5 and parts[5].isdigit() else 1,
        "h2": 8, "faq": 7, "keywords": [], "urls": [],
    }
    await run_generation(update, context, params)


# ─── МАСТЕР /new — ПОШАГОВЫЙ ДИАЛОГ ─────────────────────────────────────────
async def cmd_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "🚀 *Новый текст — шаг 1/10*\n\n"
        "Введи *бренд или название игры:*\n"
        "Пример: `GGBet` или `Chicken Road`",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    return STATE_BRAND

async def step_brand(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['brand'] = update.message.text.strip()
    geo_kb = [["DE", "IT", "LV"], ["AT", "ES", "PL"], ["UA", "CA", "NZ"], ["FI", "Другое"]]
    await update.message.reply_text(
        "📍 *Шаг 2/10 — Выбери гео:*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(geo_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_GEO

async def step_geo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    geo = update.message.text.strip().upper()
    context.user_data['geo'] = geo
    context.user_data['geo_name'] = update.message.text.strip()
    await update.message.reply_text(
        "🌐 *Шаг 3/10 — Язык текста:*\n\n"
        "Напиши язык, например:\n"
        "`Latviski` · `Deutsch` · `Italiano` · `English` · `Español`",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    return STATE_LANG

async def step_lang(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['lang'] = update.message.text.strip()
    pt_kb = [[pt] for pt in PAGE_TYPES]
    await update.message.reply_text(
        "📄 *Шаг 4/10 — Тип страницы:*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(pt_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_PAGE_TYPE

async def step_page_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['page_type'] = update.message.text.strip()
    words_kb = [["2000-2500", "2500-3000"], ["3000-3500", "3500-4000"], ["4000-4500", "4500+"]]
    await update.message.reply_text(
        "📝 *Шаг 5/10 — Объём текста (слов):*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(words_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_WORDS

async def step_words(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data['words'] = update.message.text.strip()
    h2_kb = [["6", "7", "8"], ["9", "10", "11"], ["12", "13", "14"]]
    await update.message.reply_text(
        "🔢 *Шаг 6/10 — Количество H2 заголовков:*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(h2_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_H2

async def step_h2(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data['h2'] = int(val) if val.isdigit() else 8
    faq_kb = [["5", "6", "7"], ["8", "9", "10"]]
    await update.message.reply_text(
        "❓ *Шаг 7/10 — Количество FAQ вопросов:*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(faq_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_FAQ

async def step_faq(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data['faq'] = int(val) if val.isdigit() else 7
    qty_kb = [["1", "2", "3"], ["4", "5"]]
    await update.message.reply_text(
        "📋 *Шаг 8/10 — Сколько уникальных версий текста сгенерировать?*",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(qty_kb, one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_QTY

async def step_qty(update: Update, context: ContextTypes.DEFAULT_TYPE):
    val = update.message.text.strip()
    context.user_data['qty'] = int(val) if val.isdigit() else 1
    await update.message.reply_text(
        "🔑 *Шаг 9/10 — Ключевые слова*\n\n"
        "Напиши ключевые слова — каждое с новой строки.\n"
        "Первый ключ = главный (попадёт в H1 и intro).\n\n"
        "Пример:\n"
        "```\nggbet казино\nggbet бонус\nggbet регистрация\nggbet отзывы```\n\n"
        "Или напиши *пропустить* если ключи не нужны.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["пропустить"]], one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_KEYWORDS

async def step_keywords(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() in ("пропустить", "skip", "-"):
        context.user_data['keywords'] = []
    else:
        kws = [k.strip() for k in text.split('\n') if k.strip()]
        context.user_data['keywords'] = kws
    await update.message.reply_text(
        "🔗 *Шаг 10/10 — URL конкурентов для анализа*\n\n"
        "Вставь URL конкурентов — каждый с новой строки.\n"
        "Я проанализирую их структуру, заголовки и entity map.\n\n"
        "Пример:\n"
        "```\nhttps://competitor1.com/ggbet\nhttps://competitor2.com/review\nhttps://competitor3.com```\n\n"
        "Или напиши *пропустить* — тогда сделаю live SERP анализ.",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["пропустить"]], one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_URLS

async def step_urls(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() in ("пропустить", "skip", "-"):
        context.user_data['urls'] = []
    else:
        urls = [u.strip() for u in text.split('\n') if u.strip().startswith('http')]
        context.user_data['urls'] = urls

    # Показываем сводку перед генерацией
    d = context.user_data
    kw_preview = "\n".join(f"  • {k}" for k in d.get('keywords', [])) or "  (не указаны — live SERP)"
    url_preview = "\n".join(f"  • {u}" for u in d.get('urls', [])) or "  (не указаны — live SERP)"

    summary = (
        f"✅ *Всё готово! Проверь параметры:*\n\n"
        f"🏷 Бренд: `{d.get('brand')}`\n"
        f"📍 Гео: `{d.get('geo')} — {d.get('geo_name')}`\n"
        f"🌐 Язык: `{d.get('lang')}`\n"
        f"📄 Тип: `{d.get('page_type')}`\n"
        f"📝 Слов: `{d.get('words')}`\n"
        f"🔢 H2: `{d.get('h2')}` · FAQ: `{d.get('faq')}`\n"
        f"📋 Версий: `{d.get('qty')}`\n\n"
        f"🔑 Ключи:\n{kw_preview}\n\n"
        f"🔗 URL конкурентов:\n{url_preview}\n\n"
        f"Нажми *Генерировать* или /cancel для отмены."
    )
    await update.message.reply_text(
        summary,
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup([["✅ Генерировать"], ["❌ Отмена"]], one_time_keyboard=True, resize_keyboard=True)
    )
    return STATE_CONFIRM

async def step_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if "отмен" in text.lower() or text == "❌ Отмена":
        return await cmd_cancel(update, context)

    params = dict(context.user_data)
    await run_generation(update, context, params)
    context.user_data.clear()
    return ConversationHandler.END


# ─── ГЕНЕРАЦИЯ И ОТПРАВКА ─────────────────────────────────────────────────────
async def run_generation(update: Update, context: ContextTypes.DEFAULT_TYPE, params: dict):
    qty = params.get("qty", 1)
    status_msg = await update.message.reply_text(
        f"⏳ Генерирую {qty} текст(а)...\n"
        f"Бренд: {params.get('brand')} · Гео: {params.get('geo')}\n"
        f"Обычно 30–120 секунд.",
        reply_markup=ReplyKeyboardRemove()
    )

    try:
        markdown_text = generate_content(params)

        # Конвертируем в docx
        docx_bytes = markdown_to_docx(markdown_text)

        # Имя файла
        brand_safe = re.sub(r'[^\w]', '_', params.get('brand', 'article'))
        geo_code   = params.get('geo', 'XX')
        qty_str    = f"_x{qty}" if qty > 1 else ""
        filename   = f"{brand_safe}_{geo_code}{qty_str}.docx"

        await status_msg.delete()

        doc_io = io.BytesIO(docx_bytes)
        doc_io.name = filename

        word_count = len(markdown_text.split())
        kw_info = ""
        if params.get("keywords"):
            kw_info = f"\n🔑 Главный ключ: `{params['keywords'][0]}`"
        url_info = ""
        if params.get("urls"):
            url_info = f"\n🔗 Проанализировано URL: {len(params['urls'])}"

        await update.message.reply_document(
            document=doc_io,
            filename=filename,
            caption=(
                f"✅ *Готово!*\n"
                f"📋 Версий: {qty} · Слов: ~{word_count:,}"
                f"{kw_info}{url_info}\n"
                f"📎 `{filename}`"
            ),
            parse_mode="Markdown"
        )

    except Exception as e:
        await status_msg.edit_text(
            f"❌ Ошибка: {str(e)}\n\nПопробуй ещё раз или /cancel"
        )


# ─── ЗАПУСК ──────────────────────────────────────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN не задан!")
    if not DEEPSEEK_KEY:
        raise ValueError("DEEPSEEK_KEY не задан!")

    print("🤖 iGaming SEO Bot v4 запускается...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Conversation handler для /new
    conv = ConversationHandler(
        entry_points=[CommandHandler("new", cmd_new)],
        states={
            STATE_BRAND:     [MessageHandler(filters.TEXT & ~filters.COMMAND, step_brand)],
            STATE_GEO:       [MessageHandler(filters.TEXT & ~filters.COMMAND, step_geo)],
            STATE_LANG:      [MessageHandler(filters.TEXT & ~filters.COMMAND, step_lang)],
            STATE_PAGE_TYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, step_page_type)],
            STATE_WORDS:     [MessageHandler(filters.TEXT & ~filters.COMMAND, step_words)],
            STATE_H2:        [MessageHandler(filters.TEXT & ~filters.COMMAND, step_h2)],
            STATE_FAQ:       [MessageHandler(filters.TEXT & ~filters.COMMAND, step_faq)],
            STATE_QTY:       [MessageHandler(filters.TEXT & ~filters.COMMAND, step_qty)],
            STATE_KEYWORDS:  [MessageHandler(filters.TEXT & ~filters.COMMAND, step_keywords)],
            STATE_URLS:      [MessageHandler(filters.TEXT & ~filters.COMMAND, step_urls)],
            STATE_CONFIRM:   [MessageHandler(filters.TEXT & ~filters.COMMAND, step_confirm)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        allow_reentry=True,
    )

    app.add_handler(CommandHandler("start",  cmd_start))
    app.add_handler(CommandHandler("quick",  cmd_quick))
    app.add_handler(CommandHandler("types",  cmd_types))
    app.add_handler(conv)
    # Быстрый режим — любое текстовое сообщение вне диалога
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_quick))

    print("✅ Бот работает. Ctrl+C для остановки.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
