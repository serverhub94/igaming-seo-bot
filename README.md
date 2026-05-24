# iGaming SEO Telegram Bot — Гайд по установке

## Что делает бот
Пишешь в Telegram: `GGBet, LV, латышский, Casino Review, 3000 слов`
Бот генерирует полный SEO-текст по мастер-промту v3 и присылает обратно.

---

## Шаг 1 — Создай Telegram-бота (2 минуты)

1. Открой Telegram, найди **@BotFather**
2. Напиши `/newbot`
3. Придумай имя бота (например: `iGaming SEO Writer`)
4. Придумай username (например: `igaming_seo_bot`) — должен заканчиваться на `bot`
5. BotFather пришлёт **токен** вида: `7123456789:AAHdqTcvCH1vGWJxfSeofSs0K67UKCX4fQ4`
6. **Сохрани этот токен** — он нужен в следующих шагах

---

## Шаг 2 — Получи Anthropic API ключ

1. Зайди на **console.anthropic.com**
2. Зарегистрируйся / войди
3. Слева: **API Keys** → **Create Key**
4. Скопируй ключ вида: `sk-ant-api03-...`
5. **Сохрани** — он показывается только один раз

> Стоимость: ~$0.003 за 1000 токенов (один текст 3000 слов ≈ $0.05–0.10)

---

## Шаг 3 — Загрузи код на Railway (бесплатный хостинг)

### Вариант A — через GitHub (рекомендую)

1. Создай аккаунт на **github.com**
2. Создай новый репозиторий (например: `igaming-seo-bot`)
3. Загрузи файлы из этой папки:
   - `bot.py`
   - `requirements.txt`
   - `railway.toml`
   - `.gitignore`
4. Зайди на **railway.app** → войди через GitHub
5. Нажми **New Project** → **Deploy from GitHub repo**
6. Выбери свой репозиторий → Railway автоматически задеплоит

### Вариант B — через Railway CLI

```bash
npm install -g @railway/cli
railway login
railway init
railway up
```

---

## Шаг 4 — Добавь переменные окружения в Railway

1. В Railway открой свой проект
2. Перейди в **Variables**
3. Добавь две переменные:

```
TELEGRAM_TOKEN = твой_токен_от_BotFather
ANTHROPIC_KEY  = твой_ключ_от_Anthropic
```

4. Railway автоматически перезапустит бота

---

## Шаг 5 — Проверь что бот работает

1. Найди своего бота в Telegram по username
2. Напиши `/start`
3. Бот ответит приветствием
4. Напиши тестовый запрос:
   ```
   GGBet, LV, латышский, Casino Review, 3000 слов
   ```
5. Жди 30–90 секунд → получи готовый текст

---

## Как писать запросы боту

### Минимальный формат:
```
GGBet, Латвия, латышский
```

### Полный формат:
```
GGBet, LV, Latviski, Casino Review, 3500 слов, 10 H2, 7 FAQ
```

### Примеры запросов:
```
Chicken Road, IT, Italiano, Game Review, 3000 слов
Vbet, DE, Deutsch, Bonus Page, 2500 слов
1xBet, PL, Polski, Casino Review
Aviator, ES, Español, Game Review, 4000 слов, 12 H2, 8 FAQ
```

### Поддерживаемые гео (авто-определение лицензий и RG):
| Код | Страна | Лицензия |
|-----|--------|----------|
| DE | Германия | GGL |
| IT | Италия | ADM |
| LV | Латвия | IAUI / Curaçao |
| AT | Австрия | Austrian Gaming Authority |
| ES | Испания | DGOJ |
| PL | Польша | MF |
| UA | Украина | KRAIL |
| CA | Канада | Kahnawake |

---

## Ограничение доступа (опционально)

Чтобы бот отвечал только тебе:

1. Найди свой Telegram user_id через **@userinfobot**
2. В файле `bot.py` найди строку:
   ```python
   ALLOWED_USER_IDS = []
   ```
3. Замени на:
   ```python
   ALLOWED_USER_IDS = [123456789]  # твой user_id
   ```
4. Перезагрузи бота на Railway

---

## Стоимость

| Сервис | Стоимость |
|--------|-----------|
| Railway | $5/месяц (Hobby план) или бесплатно с лимитами |
| Anthropic API | ~$0.05–0.15 за один текст 3000 слов |
| Telegram Bot API | Бесплатно |

При 50 текстах в месяц: ~$7–12 на API + $5 Railway = **~$12–17/месяц**

---

## Возможные проблемы

**Бот не отвечает:**
- Проверь TELEGRAM_TOKEN в Railway Variables
- Проверь логи: Railway → твой проект → Deployments → View Logs

**Ошибка Claude API:**
- Проверь ANTHROPIC_KEY
- Проверь баланс на console.anthropic.com

**Текст обрывается:**
- Это нормально — Telegram ограничивает 4096 символов на сообщение
- Бот автоматически разбивает текст на части

---

## Локальный запуск (для тестирования)

```bash
# Установи зависимости
pip install -r requirements.txt

# Создай .env файл
cp .env.example .env
# Заполни TELEGRAM_TOKEN и ANTHROPIC_KEY в .env

# Запусти
python bot.py
```

---

*Версия: v3 | Мастер-промт: iGaming SEO v3 Full*
