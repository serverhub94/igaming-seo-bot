import os
import re
import io
import tempfile
from openai import OpenAI
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from telegram import Update
from telegram.ext import (
    Application,
    MessageHandler,
    CommandHandler,
    filters,
    ContextTypes,
)

# ─── НАСТРОЙКИ ───────────────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
DEEPSEEK_KEY   = os.environ.get("DEEPSEEK_KEY", "")

ALLOWED_USER_IDS = []  # пример: [123456789]. Пусто = доступ для всех.

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
Optimize for:
  — Intent satisfaction
  — Semantic coverage
  — Topical completeness
  — Entity relevance
  — Commercial usefulness
  — Trust signals
  — UX scanability
Exact-match keywords = final technical adjustment ONLY.
Primary goal: cover intent BETTER than SERP competitors.

━━━ SERP ANALYSIS ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Determine:
  — Which semantic entities Google considers mandatory
  — Which subtopics appear in the majority of competitors
  — Which commercial blocks are present
  — Which trust sections are used
  — Which FAQ questions genuinely close intent
  — Which answer blocks appear in Featured Snippets

━━━ SERP MEDIAN LOGIC ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Calculate MEDIAN from relevant page group only.
DO NOT mix different intent types.

━━━ TF-IDF SEMANTIC EXTRACTION ━━━━━━━━━━━━━━━━━━━━
Extract: Core entities · Related entities · Co-occurrence phrases
Commercial: bonus · withdrawal · payments · app · promo code
            support · live betting · RTP · verification · licenses
Trust: responsible gambling · KYC · MGA · Curacao
       real reviews · security · SSL · account verification

━━━ FEATURED SNIPPET OPTIMIZATION ━━━━━━━━━━━━━━━━━
In first 20-30% add block targeting:
Featured Snippets · AI Overviews · People Also Ask · Passage Ranking
Formats: definition (40-60 words) · direct answer · comparison · steps · table

━━━ CONTENT ARCHITECTURE ━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Intro: answer-first, 60-80 words
  — Each H2: closes ONE specific user intent
  — Trust blocks: MANDATORY
  — FAQ: genuinely useful only

━━━ HEADING DEPTH RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Each H2: minimum 200-350 words
  — Develop from FOUR angles: What / Why / How / Action
  — No thin sections

━━━ SECTION INTRO + VISUAL ELEMENT RULES ━━━━━━━━━━
Every H2 MUST follow:
  1. INTRO (1-2 sentences) — frame the topic, why it matters NOW
  2. VISUAL ELEMENT (mandatory, ONE per section):
     TABLE: comparing options/specs/rates/limits
     BULLET LIST (use "- " prefix): features/pros-cons/requirements
     NUMBERED LIST (use "1. " prefix): steps/process/ranking
     Min 4 items per list, min 3 rows per table
  3. BODY TEXT: expand, add context, min 2 paragraphs
  FORMAT ROTATION: never same format in two consecutive H2s

━━━ GOOGLE QUALITY RESTRICTIONS ━━━━━━━━━━━━━━━━━━━
  — Keyword density <= 2.5-3%
  — No stuffing, no AI filler, no robotic phrasing
  — Natural language, high scanability
  — Long sentences < 5%

━━━ E-E-A-T REQUIREMENTS ━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Expert tone · Trust signals · License references
  — Transparent pros/cons · Responsible gambling
  — Author bio + date at end

━━━ WRITING STYLE ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WRITE: naturally · varied rhythm · expert but readable
DO NOT: AI filler · empty paragraphs · guaranteed wins

━━━ OUTPUT FORMAT (CRITICAL) ━━━━━━━━━━━━━━━━━━━━━━
Use EXACTLY this markdown structure — it will be converted to .docx:

# H1 Title here

Intro paragraph text here.

## H2 Section Title

Intro sentence(s) for this section.

| Column 1 | Column 2 | Column 3 |
|---|---|---|
| Row 1 | Data | Data |
| Row 2 | Data | Data |

Body text paragraphs here.

### H3 Subsection (if needed)

Text here.

For bullet lists use:
- Item one
- Item two
- Item three
- Item four

For numbered lists use:
1. Step one
2. Step two
3. Step three
4. Step four

End the article with:
## FAQ

**Question 1?**
Answer text.

**Question 2?**
Answer text.

---
*Author: [Name] — [brief bio]. Last updated: [Month Year].*
*Responsible gambling disclaimer here.*

════════════════════════════════════════════════════
HOW TO PARSE USER INPUT
════════════════════════════════════════════════════
User sends: "GGBet, LV, латышский, Casino Review, 3000 слов"
Extract: Brand · Geo · Language · Page Type · Word count (default 3000) · H2 count (default 8) · FAQ count (default 7)
Apply full system above. Generate complete article in specified language.
"""

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
}

# ─── ХЕЛПЕРЫ ─────────────────────────────────────────────────────────────────
def detect_geo(text: str) -> dict:
    text_upper = text.upper()
    for code, data in GEO_DATA.items():
        if code in text_upper:
            return {**data, "code": code}
    return {"license": "relevant local gambling regulator",
            "rg": "local responsible gambling resources", "code": "??"}


def build_user_prompt(user_text: str, geo_data: dict) -> str:
    return (
        f"{user_text}\n\n"
        f"[Auto-detected geo]\n"
        f"License: {geo_data['license']}\n"
        f"RG resources: {geo_data['rg']}"
    )


def extract_brand_and_geo(text: str) -> str:
    """Extract first two meaningful words for filename."""
    parts = [p.strip() for p in text.replace(",", " ").split() if len(p.strip()) > 1]
    name = "_".join(parts[:2]) if len(parts) >= 2 else parts[0] if parts else "article"
    return re.sub(r'[^\w\-]', '', name)


# ─── MARKDOWN → DOCX КОНВЕРТЕР ───────────────────────────────────────────────
def markdown_to_docx(markdown_text: str) -> bytes:
    """Convert markdown text to a formatted .docx file."""
    doc = Document()

    # ── Стили документа ──────────────────────────────────────────────────────
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    def style_heading(paragraph, level: int):
        sizes = {1: 22, 2: 16, 3: 13}
        paragraph.style = f'Heading {level}'
        run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
        run.font.bold = True
        run.font.size = Pt(sizes.get(level, 12))
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D) if level == 1 else RGBColor(0x2E, 0x75, 0xB6)

    def add_paragraph(text: str, bold=False, italic=False):
        if not text.strip():
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text.strip())
        run.bold = bold
        run.italic = italic
        return p

    def parse_inline(paragraph, text: str):
        """Handle **bold** and *italic* inline."""
        parts = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
            elif part.startswith('*') and part.endswith('*'):
                run = paragraph.add_run(part[1:-1])
                run.italic = True
            else:
                paragraph.add_run(part)

    def add_table(rows_data: list):
        if not rows_data or len(rows_data) < 2:
            return
        # Убираем строку-разделитель (|---|---|)
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
            for j, cell_text in enumerate(row_data):
                if j < col_count:
                    cell = row.cells[j]
                    cell.text = cell_text
                    # Заголовок таблицы
                    if i == 0:
                        for run in cell.paragraphs[0].runs:
                            run.bold = True
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                        cell.paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        # Синий фон заголовка
                        from docx.oxml.ns import qn
                        from docx.oxml import OxmlElement
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        shd = OxmlElement('w:shd')
                        shd.set(qn('w:val'), 'clear')
                        shd.set(qn('w:color'), 'auto')
                        shd.set(qn('w:fill'), '2E75B6')
                        tcPr.append(shd)

        doc.add_paragraph()

    # ── Парсинг markdown ──────────────────────────────────────────────────────
    lines = markdown_text.split('\n')
    i = 0
    in_bullet = False
    in_numbered = False

    while i < len(lines):
        line = lines[i]

        # Горизонтальная линия
        if re.match(r'^---+$', line.strip()):
            doc.add_paragraph('─' * 60)
            i += 1
            continue

        # H1
        if line.startswith('# ') and not line.startswith('## '):
            p = doc.add_heading(line[2:].strip(), level=1)
            style_heading(p, 1)
            in_bullet = in_numbered = False
            i += 1
            continue

        # H2
        if line.startswith('## ') and not line.startswith('### '):
            p = doc.add_heading(line[3:].strip(), level=2)
            style_heading(p, 2)
            in_bullet = in_numbered = False
            i += 1
            continue

        # H3
        if line.startswith('### '):
            p = doc.add_heading(line[4:].strip(), level=3)
            style_heading(p, 3)
            in_bullet = in_numbered = False
            i += 1
            continue

        # Таблица — собираем все строки таблицы
        if line.startswith('|'):
            table_rows = []
            while i < len(lines) and lines[i].startswith('|'):
                table_rows.append(lines[i])
                i += 1
            add_table(table_rows)
            continue

        # Маркированный список
        if re.match(r'^[-*•] ', line):
            item_text = re.sub(r'^[-*•] ', '', line).strip()
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_after = Pt(2)
            parse_inline(p, item_text)
            in_bullet = True
            in_numbered = False
            i += 1
            continue

        # Нумерованный список
        if re.match(r'^\d+\. ', line):
            item_text = re.sub(r'^\d+\. ', '', line).strip()
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_after = Pt(2)
            parse_inline(p, item_text)
            in_numbered = True
            in_bullet = False
            i += 1
            continue

        # Пустая строка
        if not line.strip():
            in_bullet = in_numbered = False
            i += 1
            continue

        # Жирный абзац (FAQ вопросы **Question?**)
        if line.startswith('**') and line.endswith('**'):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(line.strip('*').strip())
            run.bold = True
            i += 1
            continue

        # Курсив / обычный текст (итальянские/дисклеймер строки)
        if line.startswith('*') and line.endswith('*'):
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(line.strip('*').strip())
            run.italic = True
            run.font.size = Pt(10)
            i += 1
            continue

        # Обычный абзац с inline форматированием
        in_bullet = in_numbered = False
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        parse_inline(p, line.strip())
        i += 1

    # ── Сохраняем в bytes ─────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ─── КОМАНДЫ ─────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *iGaming SEO Content Bot* готов!\n\n"
        "Пишешь параметры → получаешь готовый *.docx* файл.\n\n"
        "*Формат:*\n"
        "`Бренд, Гео, Язык, Тип страницы, Слова`\n\n"
        "*Примеры:*\n"
        "• `GGBet, LV, латышский, Casino Review, 3000 слов`\n"
        "• `Chicken Road, IT, Italiano, Game Review`\n"
        "• `Vbet, DE, Deutsch, Bonus Page, 2500 слов`\n\n"
        "⏳ Генерация: 30–90 секунд → получишь .docx файл\n\n"
        "/help — помощь",
        parse_mode="Markdown"
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📋 *Как использовать:*\n\n"
        "*Минимум:* `GGBet, Латвия, латышский`\n\n"
        "*Полный:* `GGBet, LV, Latviski, Casino Review, 3500 слов, 10 H2, 7 FAQ`\n\n"
        "*Типы страниц:*\n"
        "Casino Review · Game Review · Bonus Page\n"
        "Sportsbook Review · How to Play · Comparison\n\n"
        "*Гео:* DE · IT · LV · AT · ES · PL · UA · CA\n\n"
        "📎 Результат — готовый .docx файл",
        parse_mode="Markdown"
    )


# ─── ОСНОВНОЙ ОБРАБОТЧИК ─────────────────────────────────────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    if ALLOWED_USER_IDS and user.id not in ALLOWED_USER_IDS:
        await update.message.reply_text("⛔ Доступ запрещён.")
        return

    user_text = update.message.text.strip()
    if not user_text:
        return

    status_msg = await update.message.reply_text(
        "⏳ Генерирую текст и создаю .docx...\n"
        "Обычно 30–90 секунд."
    )

    try:
        geo_data = detect_geo(user_text)
        enriched_prompt = build_user_prompt(user_text, geo_data)

        # ── DeepSeek API ──────────────────────────────────────────────────────
        client = OpenAI(
            api_key=DEEPSEEK_KEY,
            base_url="https://api.deepseek.com"
        )

        response = client.chat.completions.create(
            model="deepseek-v4-pro",
            max_tokens=20000,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": enriched_prompt}
            ]
        )

        markdown_text = response.choices[0].message.content

        # ── Конвертация в .docx ───────────────────────────────────────────────
        docx_bytes = markdown_to_docx(markdown_text)

        # ── Имя файла ─────────────────────────────────────────────────────────
        filename_base = extract_brand_and_geo(user_text)
        geo_info = f"_{geo_data['code']}" if geo_data["code"] != "??" else ""
        filename = f"{filename_base}{geo_info}.docx"

        # ── Отправляем файл ───────────────────────────────────────────────────
        await status_msg.delete()

        doc_file = io.BytesIO(docx_bytes)
        doc_file.name = filename

        word_count = len(markdown_text.split())
        await update.message.reply_document(
            document=doc_file,
            filename=filename,
            caption=(
                f"✅ *Готово{geo_info.replace('_', ' ')}*\n"
                f"📄 Слов: ~{word_count:,} · Файл: {filename}"
            ),
            parse_mode="Markdown"
        )

    except Exception as e:
        await status_msg.edit_text(
            f"❌ Ошибка: {str(e)}\n\n"
            "Попробуй ещё раз или напиши /start"
        )


# ─── ЗАПУСК ──────────────────────────────────────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN не задан!")
    if not DEEPSEEK_KEY:
        raise ValueError("DEEPSEEK_KEY не задан!")

    print("🤖 iGaming SEO Bot (DeepSeek + .docx) запускается...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help",  cmd_help))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("✅ Бот работает. Ctrl+C для остановки.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
