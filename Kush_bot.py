from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from dotenv import load_dotenv
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ConversationHandler,
    MessageHandler,
    filters,
    ContextTypes
)
import openpyxl
import datetime
import os


load_dotenv()

TOKEN = os.getenv("TOKEN")
FILE_PATH = "/root/kush_bot/leads.xlsx"

# Этапы
STEP_1, STEP_2, STEP_3, STEP_4, STEP_5, STEP_6, STEP_7, STEP_8, STEP_9 = range(9)
user_data = {}

# === Функция создания Excel файла, если его нет ===
def ensure_excel_exists():
    if not os.path.exists(FILE_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Лиды"
        ws.append([
            "Имя пользователя",
            "Контактная информация",
            "Бизнес",
            "Крепость бизнеса",
            "Послевкусие",
            "Секретный ингредиент",
            "Дата и время"
        ])
        wb.save(FILE_PATH)

# === /start ===
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("🍸 Конечно, наливай!", callback_data="start_game")],
        [InlineKeyboardButton("🤔 Хочу узнать, что за игра", callback_data="info")]
    ]

    # путь к локальному фото
    photo_path = os.path.join(os.path.dirname(__file__), "foto", "5.jpg")  # или "1.png" — смотри по названию

    try:
        with open(photo_path, "rb") as photo:
            await update.message.reply_photo(
                photo=photo,
                caption=(
                    "🥁 Добро пожаловать в KUSH-Бар бизнеса — место, где мы смешиваем идеальные рецепты брендов.\n\n"
                    "Я — Бармен. Сегодня я приготовлю коктейль из твоего бизнеса.\n\n"
                    "🚀 5 минут — и ты узнаешь:\n"
                    "• какой у тебя “вкус бизнеса;”\n"
                    "• что делает его горьким или пресным.\n"
                    "А также получишь свой фирменный рецепт 💎\n\n"
                    "Готов?"
                ),
                reply_markup=InlineKeyboardMarkup(keyboard)
            )
    except FileNotFoundError:
        await update.message.reply_text("⚠️ Ошибка: не найдено изображение для старта.")
        print(f"Файл не найден: {photo_path}")

    return STEP_1

# === 1. Кнопки ===
async def step1_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()  # закрывает "часики" Telegram

    if query.data == "info":
        await query.message.reply_text(
            "🎯 Это бизнес-игра, где ты узнаешь свой 'вкус бренда'\n\n"
            "Пройди мини-квиз и получи 'рецепт' своего бизнеса 🍸",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🍸 Конечно, наливай!", callback_data="start_game")]
            ])
        )
        return STEP_1

    if query.data == "start_game":
        keyboard = [
            [InlineKeyboardButton("🍹 Перейти в канал", url="https://t.me/KUSHpodklush")],
            [InlineKeyboardButton("✅ Я в клубе", callback_data="in_club")]
        ]

        await query.message.reply_text(
            "Чтобы попасть за стойку бара, нужно показать, что ты в 'клубе своих' 😉\n\n"
            "Вступай в наш Telegram-канал — там мы делимся рецептами успеха:\n"
            "📜 шаблоны, кейсы, гайды без воды.\n"
            "После этого нажми ✅ 'Я в клубе'",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return STEP_2



# === 2. Подтверждение ===
async def step2_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "in_club":
        user_id = query.from_user.id
        channel_id = "@KUSHpodklush"  # Лучше заменить на числовой ID, если приватный

        try:
            member = await context.bot.get_chat_member(channel_id, user_id)
            status = member.status

            if status in ["member", "administrator", "creator"]:
                await query.edit_message_text(
                    "Отлично! 🎉\n"
                    "Расскажи, какой у тебя основной ингредиент бизнеса — чем ты занимаешься?"
                )
                return STEP_3
            else:
                await query.edit_message_text(
                    "❌ Похоже, ты ещё не подписан на канал.\n"
                    "Пожалуйста, подпишись и нажми кнопку снова 👇",
                    reply_markup=InlineKeyboardMarkup([
                        [InlineKeyboardButton("🍹 Перейти в канал", url="https://t.me/KUSHpodklush")],
                        [InlineKeyboardButton("✅ Я в клубе", callback_data="in_club")]
                    ]))
                return STEP_2

        except Exception as e:
            print("Ошибка при проверке подписки:", e)
            await query.edit_message_text(
                "⚠️ Не удалось проверить подписку. Возможно, бот не является администратором канала.\n\n"
                "Попробуй снова или подпишись вручную 👉 https://t.me/KUSHpodklush"
            )
            return STEP_2

# === 3. Вид бизнеса ===
async def handle_business_type(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data[user_id] = {"business": update.message.text}

    keyboard = [
        [InlineKeyboardButton("🐣 Только начал — хочу поставить на поток", callback_data="1_var")],
        [InlineKeyboardButton("💼 Работаю стабильно, но всё вручную", callback_data="2_var")],
        [InlineKeyboardButton("🏢 Есть команда и клиенты, но нужна система", callback_data="3_var")]
    ]
    await update.message.reply_text(
        "Выбирай — какая крепость у твоего бизнеса? 🍷",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return STEP_4

# === 4. Крепость ===
async def step4_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    user_data[user_id]["strength"] = query.data

    keyboard = [
        [InlineKeyboardButton("😅 Лёгкий хаос и миллион задач", callback_data="1_var")],
        [InlineKeyboardButton("😬 Всё вроде работает, но как-то неэффективно", callback_data="2_var")],
        [InlineKeyboardButton("💪 Чётко, но не хватает визуала и стиля", callback_data="3_var")],
        [InlineKeyboardButton("😎 Всё под контролем — просто интересно, что скажет бармен", callback_data="4_var")]
    ]
    await query.edit_message_text(
        "Когда ты думаешь про свой бизнес, что чувствуешь чаще всего?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return STEP_5

# === 5. Послевкусие ===
async def step5_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    user_data[user_id]["feeling"] = query.data

    keyboard = [
        [InlineKeyboardButton("⚖️ Юридическую защиту — не хочу рисковать", callback_data="1_var")],
        [InlineKeyboardButton("🎨 Красивый визуал и сайт", callback_data="2_var")],
        [InlineKeyboardButton("🤖 Автоматизацию — чтобы заявки не терялись", callback_data="3_var")],
        [InlineKeyboardButton("🧭 Всё вместе: хочу, чтобы бизнес работал сам", callback_data="4_var")]
    ]
    await query.edit_message_text(
        "Окей, теперь важный вопрос:\n"
        "Что ты хотел(а) бы добавить в свой “бизнес-коктейль”?",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
    return STEP_6

# === 6. Анализ “рецепта” ===
async def step6_buttons(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id
    user_data[user_id]["secret"] = query.data

    await query.edit_message_text(
        "Подожди, я мешаю… \n\n"
        "🍸 Слышны звуки льда, бармен встряхивает шейкер\n"
        "💭 Так-так… вижу лёгкий оттенок хаоса, нотки усталости и щепотку перфекционизма.\n\n"
        "Кажется, у тебя бизнес с потенциалом, которому не хватает системы и визуальной подачи.\n\n"
        "Хочешь, я покажу рецепт идеального коктейля для твоего бизнеса?",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🔥 Да, покажи рецепт!", callback_data="show_recipe")]
        ])
    )
    return STEP_7

# === 7. Рецепт бизнеса (универсальная сборка по комбинации ответов) ===
async def show_recipe(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = query.from_user.id

    data = user_data.get(user_id, {})
    strength = data.get("strength", "")
    feeling = data.get("feeling", "")
    secret = data.get("secret", "")
    business = data.get("business", "Ваш бизнес")[:200]

    strength_head = {
        "1_var": "БАЗОВЫЙ СТАРТ 🍹",
        "2_var": "СТАБИЛЬНЫЙ СМУЗИ 🍸",
        "3_var": "КОМАНДНЫЙ ДЖИН 🍷",
        "4_var": "VIP-РЕЦЕПТ ✨"
    }
    headline = strength_head.get(strength, "Авторский коктейль KUSH 🍸")

    strength_ingredients = {
        "1_var": ["свежая идея", "щепотка энтузиазма", "немного хаоса (но это лечится)"],
        "2_var": ["уверенность", "пара ручных процессов", "немного недосистемы"],
        "3_var": ["опытная команда", "фирменный стиль", "немного организационного хаоса"],
        "4_var": ["отлаженные процессы", "брендбук", "ясная структура ответственности"]
    }
    ingredients = strength_ingredients.get(strength, ["идея", "сила воли"])

    feeling_notes = {
        "1_var": "Лёгкий хаос: приберём процессы и поставим воронку, чтобы ты перестал терять заявки.",
        "2_var": "Неэффективно: автоматизируем рутинные операции и снизим ручную работу.",
        "3_var": "Не хватает визуала: сделаем фирменный стиль и продающий лендинг.",
        "4_var": "Всё под контролем: усилим продажи и выведем систему на новый уровень."
    }
    feeling_text = feeling_notes.get(feeling, "Мы подберём набор работ под вашу ситуацию.")

    secret_help = {
        "1_var": "🧾 С юридической поддержкой — чтобы старт был безопасным.",
        "2_var": "🎨 С дизайном и визуалом — чтобы продукт начал продавать себя сам.",
        "3_var": "🤖 С автоматизацией и CRM — чтобы заявки не терялись.",
        "4_var": "🧭 С полным пакетом (юрист, дизайн, автоматизация) — всё под ключ."
    }
    secret_text = secret_help.get(secret, "Подберём оптимальный набор инструментов.")

    combo_hint = ""
    if strength == "1_var" and feeling == "3_var":
        combo_hint = "Рекомендуем начать с простого CRM + шаблонов коммуникации — это сразу уменьшит потери лидов."
    elif strength == "2_var" and feeling == "2_var":
        combo_hint = "Оптимизация процессов и внедрение CRM дадут эффект уже через 7–14 дней."
    elif strength == "3_var" and secret == "2_var":
        combo_hint = "Команде достаточно нового фирменного гайдлайна и шаблонов — визуал поднимет ценник предложений."

    recipe_lines = [
        f"Твой коктейль — «{headline}»\n",
        "Состав:",
    ]
    recipe_lines += [f"• {ing}" for ing in ingredients]
    recipe_lines.append("")
    recipe_lines.append("Наши бармены помогут:")
    recipe_lines.append(f"{secret_text}")
    recipe_lines.append(f"{feeling_text}")
    if combo_hint:
        recipe_lines.append("")
        recipe_lines.append(f"💡 Совет бармена: {combo_hint}")
    recipe_lines.append("")
    recipe_lines.append(f"Коротко про твой бизнес: {business}")
    recipe_lines.append("")
    recipe_lines.append("Всё это — возможный план действий на 14 дней.\n")
    recipe_lines.append("Хочешь получить бесплатный разбор рецепта твоего бизнеса с барменом лично?")

    recipe_text = "\n".join(recipe_lines)

    # Пути к фотографиям для каждого коктейля
    photo_paths = {
        "1_var": os.path.join("foto", "1.png"),
        "2_var": os.path.join("foto", "2.png"),
        "3_var": os.path.join("foto", "3.png"),
        "4_var": os.path.join("foto", "4.png")
    }
    photo_path = photo_paths.get(strength)

    keyboard = [
        [InlineKeyboardButton("🎁 Да, хочу разбор", callback_data="yes_recipe")],
        [InlineKeyboardButton("⏳ Не сейчас", callback_data="later")]
    ]

    try:
        if photo_path and os.path.exists(photo_path):
            with open(photo_path, "rb") as photo:
                await query.message.reply_photo(photo=photo, caption=recipe_text,
                                                reply_markup=InlineKeyboardMarkup(keyboard))
        else:
            await query.message.reply_text(recipe_text, reply_markup=InlineKeyboardMarkup(keyboard))
    except Exception as e:
        print(f"Ошибка при отправке рецепта с фото: {e}")

    return STEP_8

# === 8. Обработка кнопок "Да, хочу рецепт" / "Не сейчас" ===
async def step8_collect(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if query.data == "later":
        await query.edit_message_text("🍷 Хорошо, приходи позже — бар всегда открыт 😉")
        return ConversationHandler.END

    # Если пользователь согласился — просто отправляем новое сообщение (а не редактируем старое)
    await query.message.reply_text(
        "Отлично! ✍️\n\n"
        "Напиши, пожалуйста:\n"
        "1️⃣ Имя\n"
        "2️⃣ Ссылку на Telegram / Instagram\n"
        "3️⃣ Чем занимаешься\n\n"
        "После этого я запишу тебя на бесплатную консультацию от команды KUSH 🍸"
    )
    return STEP_9

# === 9. Сбор текста и сохранение в Excel ===
# === 9. Сбор текста и сохранение в Excel ===
async def save_to_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    user_data[user_id]["contact_info"] = update.message.text
    data = user_data[user_id]
    ensure_excel_exists()
    try:
        wb = openpyxl.load_workbook(FILE_PATH)
        ws = wb.active
        ws.append([
            update.effective_user.first_name,
            data.get("contact_info", ""),
            data.get("business", ""),
            data.get("strength", ""),
            data.get("feeling", ""),
            data.get("secret", ""),
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M") ])
        wb.save(FILE_PATH)
    except PermissionError:
        # создаём резервную копию, если основной файл заблокирован Excel
        alt_path = FILE_PATH.replace(".xlsx", f"_{datetime.datetime.now().strftime('%H%M%S')}.xlsx")
        wb.save(alt_path)
        print(f"⚠️ Основной файл был заблокирован. Данные сохранены в {alt_path}")
    await update.message.reply_text( "✅ Готово!\n\n" "Ты официально в списке гостей KUSH-Барменов 🍸\n\n" "💬 В ближайшее время (Пн-пт 9:00-19:00 по мск) с тобой свяжется Анна, чтобы назначить встречу!" )
    return ConversationHandler.END

def main():
    app = ApplicationBuilder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            STEP_1: [CallbackQueryHandler(step1_buttons)],
            STEP_2: [CallbackQueryHandler(step2_buttons)],
            STEP_3: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_business_type)],
            STEP_4: [CallbackQueryHandler(step4_buttons)],
            STEP_5: [CallbackQueryHandler(step5_buttons)],
            STEP_6: [CallbackQueryHandler(step6_buttons)],
            STEP_7: [CallbackQueryHandler(show_recipe, pattern="show_recipe")],
            STEP_8: [CallbackQueryHandler(step8_collect, pattern="yes_recipe|later")],
            STEP_9: [MessageHandler(filters.TEXT & ~filters.COMMAND, save_to_excel)],
        },
        fallbacks=[]
    )
    app.add_handler(conv)
    print("✅ Бот запущен...")
    app.run_polling()

if __name__ == "__main__":
    main()
