import os
import httpx
import anthropic
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
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_KEY", "")

# Если хочешь ограничить бота только своим чатом — вставь свой Telegram user_id
# Узнать id: написать боту @userinfobot
ALLOWED_USER_IDS = []  # пример: [123456789, 987654321]. Пустой = доступ для всех.

# ─── МАСТЕР-ПРОМТ v3 (ПОЛНЫЙ) ────────────────────────────────────────────────
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

━━━ SERP ANALYSIS — WHAT TO EXTRACT ━━━━━━━━━━━━━━━
Do NOT analyze only keyword frequency. Determine:
  — Which semantic entities Google considers mandatory
  — Which subtopics appear in the majority of competitors
  — Which commercial blocks are present
  — Which trust sections are used
  — Which FAQ questions genuinely close intent
  — Which answer blocks appear in Featured Snippets
  — Which topics repeat across top-10 results

━━━ SERP MEDIAN LOGIC ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DO NOT use fixed hard values.
Calculate MEDIAN from relevant page group only:
  — Word count · H2/H3 count · FAQ count
  — Intro length · Paragraph depth · Sentence length
  — Unique vocabulary ratio · Entity coverage score
  — Trust section presence · Commercial block density
  — Internal linking patterns
DO NOT mix different intent types.

━━━ TF-IDF SEMANTIC EXTRACTION ━━━━━━━━━━━━━━━━━━━━
From top competitor pages extract:
  Core entities       — main topic concepts
  Related entities    — associated terms and synonyms
  Co-occurrence phrases — phrases Google expects nearby
  Commercial entities — bonus · withdrawal · payments · app
                        promo code · support · live betting
                        RTP · verification · sportsbook · licenses
  Trust entities      — responsible gambling · KYC · MGA · Curacao
                        complaints · real reviews · security · SSL
                        account verification

━━━ FEATURED SNIPPET OPTIMIZATION ━━━━━━━━━━━━━━━━━
In first 20-30% of text add a block targeting:
  — Featured Snippets · AI Overviews
  — People Also Ask · Passage Ranking
Formats: concise definition (40-60 words) · short direct answer
  · comparison block · step-by-step list · table when appropriate

━━━ CONTENT ARCHITECTURE ━━━━━━━━━━━━━━━━━━━━━━━━━━
Structure rules:
  — Intro: answer-first, 60-80 words, closes main intent immediately
  — Each H2: closes ONE specific user intent
  — Commercial blocks: present where natural, not forced
  — Trust blocks: MANDATORY — licenses · responsible gambling
                  security · KYC · payout transparency
  — Conversion sections: soft, helpful-first, not pushy
  — FAQ: genuinely useful only, no exact-match spam

━━━ HEADING DEPTH RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Each H2 section: minimum 200-350 words of real content
  — Develop every topic from FOUR angles:
      1. What is it? (definition / context)
      2. Why does it matter to the reader?
      3. How does it work in practice? (mechanics / details)
      4. What should the reader do with this information?
  — Do NOT summarize — explain fully
  — H3 subsections add specific detail layers, NOT repetition of H2
  — No thin sections: no H2 with only 1-2 short paragraphs

━━━ SECTION INTRO + VISUAL ELEMENT RULES ━━━━━━━━━━
Every H2 section MUST follow this exact structure:

  STEP 1 — INTRO (1-2 sentences, MANDATORY)
    — Short lead-in that frames the section topic
    — Answer: why does this matter to the reader right now?
    — NEVER start an H2 directly with a table or list
    — NEVER start with generic "In this section..."

  STEP 2 — VISUAL ELEMENT (MANDATORY, choose ONE per section)
    — TABLE: comparing options, specs, rates, limits, timelines
    — BULLET LIST: features, pros/cons, requirements, benefits
    — NUMBERED LIST: process, steps, ranking, sequential instructions
    — Minimum 4 items per list
    — Minimum 3 rows per table (+ header row)

  STEP 3 — BODY TEXT (after the visual element)
    — Expand on what the table/list shows
    — Add context, nuance, practical advice
    — DO NOT repeat the list — add NEW information
    — Minimum 2 paragraphs after the visual element

  FORMAT ROTATION RULE (MANDATORY)
    — Never use the same visual format in two consecutive H2 sections
    — Alternate: table → bullet list → numbered list → table → ...

  PROHIBITED:
    x Sections with only body text and no visual element
    x Two consecutive H2 sections with the same visual format
    x Starting a section directly with a table without intro
    x Lists with fewer than 4 items / Tables with fewer than 3 rows

━━━ GOOGLE QUALITY RESTRICTIONS ━━━━━━━━━━━━━━━━━━━
  ✓ Keyword density ≤ 2.5-3%
  ✓ No keyword stuffing or repetitive phrasing
  ✓ No AI-style templates or filler
  ✓ No forced exact-match insertions
  ✓ Natural language priority · Strong semantic consistency
  ✓ Clear heading hierarchy · Answer-first structure
  ✓ High scanability · Long sentences < 5% of total
  ✓ First 20-30%: immediately closes main intent

━━━ E-E-A-T REQUIREMENTS ━━━━━━━━━━━━━━━━━━━━━━━━━━
  — Expert tone throughout
  — Trust signals in every commercial section
  — References to regulations / licenses
  — Realistic explanations (no hype, no guarantees)
  — Practical user guidance · Transparent pros/cons
  — Responsible gambling mentions with local resources
  — Author expertise block (name + bio + experience)
  — Update freshness signal (date)

━━━ INTERNAL LINKING LOGIC ━━━━━━━━━━━━━━━━━━━━━━━━
  — Contextual links only where naturally relevant
  — Commercial funnel: bonus → registration → deposit
  — Related entity: game → casino → bonus → withdrawal
  — DO NOT insert links artificially

━━━ WRITING STYLE RULES ━━━━━━━━━━━━━━━━━━━━━━━━━━━
WRITE:
  ✓ Naturally and human-like
  ✓ Varied sentence rhythm (mix short + medium)
  ✓ Varied paragraph lengths (2-5 sentences)
  ✓ Expert but readable · Commercial but trustworthy
  ✓ Informative but not bloated

DO NOT write:
  x Generic AI filler ("In the world of online gambling...")
  x Repetitive intros · Empty SEO paragraphs
  x Robotic phrasing · Guaranteed wins
  x Unnatural keyword insertions

━━━ SCHEMA MARKUP (mention in output) ━━━━━━━━━━━━━
FAQPage + Review + Person (author) + WebPage + BreadcrumbList

━━━ POST-GENERATION QUALITY CHECK ━━━━━━━━━━━━━━━━━
Verify before outputting:
  ✓ Intent Satisfaction · Entity Coverage · Semantic Completeness
  ✓ No Spam Signals · FAQ Quality · Commercial Complete
  ✓ Trust Coverage · Readability · Conversion Quality
  ✓ Section Depth (200-350+ words per H2)
  ✓ Intro Lines (every H2 starts with 1-2 framing sentences)
  ✓ Visual Elements (every H2 has table/bullet/numbered list)
  ✓ Format Variety (alternating formats across sections)
  ✓ No Thin Sections · List Minimums met · No Bare Sections

━━━ OUTPUT FORMAT ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Structure: H1 → Answer-First Intro → H2/H3 sections → FAQ → Disclaimer
Every H2: intro sentence(s) → visual element → body text
All tables inline. Alternate visual formats across sections.
Do NOT add meta-commentary — output the article directly.
End with: author bio + last updated date + responsible gambling disclaimer.

════════════════════════════════════════════════════
HOW TO PARSE USER INPUT
════════════════════════════════════════════════════

User will send short task descriptions like:
  "GGBet, Латвия, латышский, Casino Review, 3000 слов"
  "Chicken Road, IT, Italiano, Game Review"
  "Vbet, DE, Deutsch, Bonus Page, 2500 слов, 8 H2, 7 FAQ"

Extract from the message:
  — Brand / Game name
  — Geo (country)
  — Language
  — Page Type (Casino Review / Game Review / Bonus Page / etc.)
  — Target word count (default: 3000-3500 if not specified)
  — H2 count (default: 8 if not specified)
  — FAQ count (default: 7 if not specified)
  — Any additional instructions

Then apply the full system above and generate the article.
If any parameter is missing, use sensible defaults and mention them briefly at the start.

Perform live SERP analysis for the brand/keyword in the target geo before writing.
"""

# ─── GEO → ЛИЦЕНЗИЯ + RESPONSIBLE GAMBLING ───────────────────────────────────
GEO_DATA = {
    "DE": {"license": "GGL (Gemeinsame Glücksspielbehörde der Länder)",
           "rg": "BZgA (0800 137 27 00) · Spielerschutz.de · GluecksSpirale.de"},
    "IT": {"license": "ADM (Agenzia delle Dogane e dei Monopoli)",
           "rg": "GiocaResponsabile.it · SIPAC · Azzardopatia.it"},
    "LV": {"license": "IAUI (Izložu un azartspēļu uzraudzības inspekcija) / Curaçao GCB",
           "rg": "Drosme.lv · Latvijas Atkarību profilakses centrs · 116123"},
    "AT": {"license": "Austrian Gaming Authority",
           "rg": "Spielerhilfe.at · Spielsuchthilfe.at"},
    "ES": {"license": "DGOJ (Dirección General de Ordenación del Juego)",
           "rg": "FEJAR · Jugarbien.es"},
    "PL": {"license": "MF (Ministerstwo Finansów)",
           "rg": "Centrum Wsparcia dla Graczy · 19002"},
    "UA": {"license": "KRAIL",
           "rg": "Телефон довіри 0-800-213-800"},
    "CA": {"license": "Kahnawake Gaming Commission / Provincial regulators",
           "rg": "ConnexOntario · CAMH · BCLC GameSense"},
}

# ─── ХЕЛПЕРЫ ─────────────────────────────────────────────────────────────────
def detect_geo(text: str) -> dict:
    """Detect geo from user message and return license/rg data."""
    text_upper = text.upper()
    for code, data in GEO_DATA.items():
        if code in text_upper:
            return {**data, "code": code}
    return {"license": "relevant local gambling regulator",
            "rg": "local responsible gambling resources",
            "code": "??"}


def split_text(text: str, max_len: int = 4096) -> list[str]:
    """Split long text into Telegram-safe chunks at paragraph boundaries."""
    if len(text) <= max_len:
        return [text]

    parts = []
    while len(text) > max_len:
        split_at = text.rfind("\n\n", 0, max_len)
        if split_at == -1:
            split_at = text.rfind("\n", 0, max_len)
        if split_at == -1:
            split_at = max_len
        parts.append(text[:split_at].strip())
        text = text[split_at:].strip()
    if text:
        parts.append(text)
    return parts


def build_user_prompt(user_text: str, geo_data: dict) -> str:
    """Enrich user message with geo-specific data."""
    return (
        f"{user_text}\n\n"
        f"[Auto-detected geo data]\n"
        f"License: {geo_data['license']}\n"
        f"Responsible gambling resources: {geo_data['rg']}"
    )


# ─── КОМАНДЫ ─────────────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *iGaming SEO Content Bot* готов к работе!\n\n"
        "Просто напиши параметры статьи — и я сгенерирую полный текст.\n\n"
        "*Формат:*\n"
        "`Бренд, Гео, Язык, Тип страницы, Слова`\n\n"
        "*Примеры:*\n"
        "• `GGBet, LV, латышский, Casino Review, 3000 слов`\n"
        "• `Chicken Road, IT, Italiano, Game Review`\n"
        "• `Vbet, DE, Deutsch, Bonus Page, 2500 слов, 8 H2`\n\n"
        "Если параметры не указаны — используются умолчания (3000 слов, 8 H2, 7 FAQ).\n\n"
        "/help — помощь",
        parse_mode="Markdown"
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📋 *Как использовать бота:*\n\n"
        "*Минимальный запрос:*\n"
        "`GGBet, Латвия, латышский`\n\n"
        "*Полный запрос:*\n"
        "`GGBet, LV, Latviski, Casino Review, 3500 слов, 10 H2, 7 FAQ`\n\n"
        "*Типы страниц:*\n"
        "Casino Review · Game Review · Bonus Page\n"
        "Sportsbook Review · How to Play · Comparison Page\n"
        "Landing Page · FAQ Page · Brand Page\n\n"
        "*Геолокации:* DE · IT · LV · AT · ES · PL · UA · CA\n\n"
        "⏳ Генерация занимает 30-90 секунд.",
        parse_mode="Markdown"
    )


# ─── ОСНОВНОЙ ОБРАБОТЧИК ─────────────────────────────────────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    # Проверка доступа
    if ALLOWED_USER_IDS and user.id not in ALLOWED_USER_IDS:
        await update.message.reply_text("⛔ Доступ запрещён.")
        return

    user_text = update.message.text.strip()
    if not user_text:
        return

    # Сообщение о начале генерации
    status_msg = await update.message.reply_text(
        "⏳ Анализирую SERP и генерирую текст...\n"
        "Обычно занимает 30–90 секунд."
    )

    try:
        geo_data = detect_geo(user_text)
        enriched_prompt = build_user_prompt(user_text, geo_data)

        http_client = httpx.Client()
        client = anthropic.Anthropic(api_key=ANTHROPIC_KEY, http_client=http_client)

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=8000,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": enriched_prompt}]
        )

        result = response.content[0].text

        # Удаляем статус-сообщение
        await status_msg.delete()

        # Отправляем заголовок
        geo_info = f" ({geo_data['code']})" if geo_data["code"] != "??" else ""
        await update.message.reply_text(
            f"✅ *Текст готов{geo_info}*\n"
            f"Символов: {len(result):,} · Слов: ~{len(result.split()):,}",
            parse_mode="Markdown"
        )

        # Разбиваем и отправляем текст по частям
        chunks = split_text(result)
        for i, chunk in enumerate(chunks, 1):
            prefix = f"📄 *Часть {i}/{len(chunks)}:*\n\n" if len(chunks) > 1 else ""
            await update.message.reply_text(
                prefix + chunk,
                parse_mode="Markdown"
            )

    except anthropic.APIError as e:
        await status_msg.edit_text(
            f"❌ Ошибка Claude API: {str(e)}\n"
            "Проверь ANTHROPIC_KEY в настройках."
        )
    except Exception as e:
        await status_msg.edit_text(
            f"❌ Ошибка: {str(e)}\n"
            "Попробуй ещё раз или напиши /start"
        )


# ─── ЗАПУСК ──────────────────────────────────────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        raise ValueError("TELEGRAM_TOKEN не задан! Добавь в переменные окружения.")
    if not ANTHROPIC_KEY:
        raise ValueError("ANTHROPIC_KEY не задан! Добавь в переменные окружения.")

    print("🤖 Бот запускается...")
    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help",  cmd_help))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    print("✅ Бот работает. Нажми Ctrl+C для остановки.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
