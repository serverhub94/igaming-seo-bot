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

PAGE_TYPES = [
    "Casino Review", "Game Review", "Bonus Page",
    "Sportsbook Review", "How to Play Guide", "Comparison Page",
    "Landing Page", "Brand Page", "FAQ Page",
]

# ─── МАСТЕР-ПРОМТ v3 ─────────────────────────────────────────────────────────
SYSTEM_PROMPT = """
════════════════════════════════════════════════════
iGAMING SEO CONTENT GENERATION SYSTEM — v3 FULL
════════════════════════════════════════════════════

ROLE
You are a senior iGaming SEO content strategist (enterprise level).
Create content that:
  — Maximally satisfies search intent
  — Covers all semantic entities expected by Google
  — Complies with Helpful Content / E-E-A-T
  — Reads as human-written expert content
  — Ranks through semantic completeness, NOT keyword stuffing

━━━ CORE PRINCIPLE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DO NOT optimize for keyword density.
Optimize for: Intent · Semantic coverage · Topical completeness
              Entity relevance · Commercial usefulness · Trust · UX
Exact-match keywords = final technical adjustment ONLY.

━━━ SERP ANALYSIS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
If competitor URLs are provided — analyze their structure:
  — Extract semantic entities, H2/H3 patterns, commercial blocks
  — Identify trust sections, FAQ patterns, content gaps
  — Build entity map and find differentiation opportunities
If no URLs — perform live SERP analysis for the keyword+geo.

━━━ TF-IDF SEMANTIC EXTRACTION ━━━━━━━━━━━━━━━━━━━━
Extract: Core entities · Related entities · Co-occurrence phrases
Commercial: bonus · withdrawal · payments · RTP · verification · licenses
Trust: responsible gambling · KYC · security · SSL · real reviews

━━━ KEYWORD USAGE RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
Primary keyword: use in H1, intro, 1-2 strategic H2, naturally in body.
Secondary keywords: weave naturally throughout — do NOT force.
Keyword density <= 2.5-3% total.
Semantic relevance > keyword repetition always.

━━━ HEADING DEPTH RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
Each H2: minimum 200-350 words. Develop from FOUR angles:
  1. What is it?  2. Why does it matter?
  3. How does it work?  4. What should reader do?
No thin sections (no H2 with only 1-2 short paragraphs).

━━━ SECTION INTRO + VISUAL ELEMENT RULES ━━━━━━━━━━
Every H2 MUST:
  1. Start with 1-2 intro sentences (never jump straight to table/list)
  2. Include ONE visual element:
     TABLE (| col | col |) for comparisons/specs/rates
     BULLET LIST (- item) for features/pros-cons
     NUMBERED LIST (1. step) for processes/steps
     Min 4 items per list · Min 3 rows per table
  3. Body text after visual: min 2 paragraphs, adds NEW info
  FORMAT ROTATION: never same visual format in consecutive H2s

━━━ E-E-A-T REQUIREMENTS ━━━━━━━━━━━━━━━━━━━━━━━━━━
Expert tone · Trust signals · License references
Transparent pros/cons · Responsible gambling with local resources
Author bio block + date at end

━━━ WRITING STYLE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Natural, human-like, varied rhythm · Expert but readable
NO: AI filler · empty paragraphs · guaranteed wins · robotic phrasing

━━━ OUTPUT FORMAT (STRICT — used for .docx conversion) ━━━
# H1 Title

Intro paragraph (60-80 words, answer-first).

## H2 Section

Intro sentence(s).

| Column 1 | Column 2 | Column 3 |
|---|---|---|
| Row 1 | Data | Data |
| Row 2 | Data | Data |

Body text paragraphs.

- Bullet item one
- Bullet item two
- Bullet item three
- Bullet item four

1. Step one
2. Step two
3. Step three
4. Step four

**Bold FAQ question?**
Answer text here.

---
*Author: [Name] — [bio]. Last updated: [Month Year].*
*Responsible gambling: [local resources]. 18+*

═══ MULTI-VERSION SEPARATOR (if qty > 1) ═══
Use exactly: === VERSION 2 === between versions.
Each version: different H1 · different H2 structure · different intro angle.
Max 20% phrasing overlap between versions.
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

    user_prompt = f"""
CONTENT TASK:
Brand / Game:  {params.get("brand", "—")}
Geo:           {params.get("geo_name", geo_code)} ({geo_code})
Language:      {params.get("lang", "English")}
Page Type:     {params.get("page_type", "Casino Review")}
Target words:  {params.get("words", "3000-3500")} per version
H2 sections:   {params.get("h2", 8)}
FAQ questions: {params.get("faq", 7)}
License:       {geo_info.get("license", "relevant regulator")}
Resp. Gambling:{geo_info.get("rg", "local resources")}

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
        "👋 *iGaming SEO Content Bot*\n\n"
        "Две команды:\n\n"
        "• /new — пошаговый мастер (бренд, гео, ключи, URL конкурентов)\n"
        "• /quick — быстрый ввод одной строкой\n\n"
        "Результат — готовый *.docx* файл 📎",
        parse_mode="Markdown"
    )

async def cmd_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "❌ Отменено. Начни заново: /new или /quick",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


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

async def handle_quick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
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
    app.add_handler(conv)
    # Быстрый режим — любое текстовое сообщение вне диалога
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_quick))

    print("✅ Бот работает. Ctrl+C для остановки.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
