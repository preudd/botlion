# -*- coding: utf-8 -*-
"""
Телеграм-бот для формирования вечернего отчёта из Excel файла с экспортом чеков.
Отправьте боту файл .xlsx - он вернёт готовый отчёт.
"""
import asyncio
import os
import sys
import tempfile
import re
from datetime import datetime

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ConversationHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

from report_parser import parse_excel_report, format_report
from rules_manager import UI_CATEGORIES, add_keyword, load_rules, remove_keyword

STATE_PICK_CATEGORY, STATE_WAIT_KEYWORD, STATE_PICK_DELETE_CATEGORY, STATE_PICK_DELETE_KEYWORD = range(4)
APP_BUILD = "rules-ui+rules-parser+prochee-v3"


def _load_bot_token() -> str | None:
    """
    Токен (по приоритету):
    1) Переменные окружения — удобно на хостинге (config.py в Git не кладём).
       Имена: REPORT_BOT_TOKEN, BOT_TOKEN, TELEGRAM_BOT_TOKEN
    2) Локальный файл config.py (копия из config.example.py)
    """
    for env_name in ('REPORT_BOT_TOKEN', 'BOT_TOKEN', 'TELEGRAM_BOT_TOKEN'):
        val = os.environ.get(env_name, '').strip()
        if val:
            return val
    try:
        from config import BOT_TOKEN as cfg_token
        t = (cfg_token or '').strip()
        if t and t != 'YOUR_BOT_TOKEN_HERE':
            return t
    except ImportError:
        pass
    return None


BOT_TOKEN = _load_bot_token() or 'YOUR_BOT_TOKEN_HERE'


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик команды /start"""
    keyboard = [
        [InlineKeyboardButton("Добавить правило", callback_data="ui:add_rule")],
        [InlineKeyboardButton("Удалить правило", callback_data="ui:delete_rule")],
        [InlineKeyboardButton("Показать правила", callback_data="ui:show_rules")],
    ]
    await update.message.reply_text(
        "Привет! Я бот для формирования вечернего отчёта.\n\n"
        "📤 Отправь мне файл Excel (.xlsx) с экспортом чеков — я сформирую отчёт и отправлю его тебе.\n\n"
        "Дата в отчёте будет взята из названия файла (например, 'Экспорт чеков от 17-01-2026.xlsx' → 17.01.2026), "
        "или сегодняшняя, если дату не удастся определить."
        , reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def ui_add_rule(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Показывает список категорий, куда можно добавить ключевое слово."""
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    keyboard: list[list[InlineKeyboardButton]] = []
    for i, cat in enumerate(UI_CATEGORIES):
        # 2 колонки
        if i % 2 == 0:
            keyboard.append([])
        keyboard[-1].append(
            InlineKeyboardButton(cat["label"], callback_data=f"cat:{cat['rule_key']}")
        )
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])

    await query.edit_message_text(
        "Выбери категорию. Потом пришли ключевое слово/фрагмент, который встречается в названии позиции.",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_CATEGORY


async def ui_choose_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    data = query.data or ""
    # cat:<rule_key>
    rule_key = data.split(":", 1)[1] if ":" in data else ""
    if not rule_key:
        await query.edit_message_text("Не удалось выбрать категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END

    context.user_data["rule_key"] = rule_key
    await query.edit_message_text(
        "Пришли ключевое слово/фрагмент.\nНапример: `аванс выпускной` или `LION`",
        reply_markup=InlineKeyboardMarkup([[InlineKeyboardButton("Отмена", callback_data="cancel")]]),
    )
    return STATE_WAIT_KEYWORD


async def ui_receive_keyword(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    keyword = (update.message.text or "").strip() if update.message else ""
    rule_key = context.user_data.get("rule_key")
    if not rule_key:
        if update.message:
            await update.message.reply_text("Не удалось определить категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END
    if not keyword:
        if update.message:
            await update.message.reply_text("Ключевое слово пустое. Попробуй ещё раз.")
        return STATE_WAIT_KEYWORD

    try:
        add_keyword(rule_key, keyword)
    except Exception as e:
        if update.message:
            await update.message.reply_text(f"Ошибка при добавлении правила: {e}")
        return STATE_WAIT_KEYWORD

    label = next((c["label"] for c in UI_CATEGORIES if c["rule_key"] == rule_key), rule_key)
    await update.message.reply_text(
        f"✅ Добавлено!\nКатегория: {label}\nКлючевое слово: {keyword}"
    )
    context.user_data.pop("rule_key", None)
    return ConversationHandler.END


async def ui_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if query:
        await query.answer()
        try:
            await query.edit_message_text("Отменено.")
        except Exception:
            # иногда сообщение уже не редактируется (если, например, уже ответили)
            pass
    context.user_data.pop("rule_key", None)
    return ConversationHandler.END


async def ui_show_rules(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    rules = load_rules()
    lines = ["Текущие правила:"]
    for cat in UI_CATEGORIES:
        rk = cat["rule_key"]
        r = rules.get(rk)
        if not r:
            continue
        lines.append(f"- {cat['label']}: " + ", ".join(r.keywords))

    await query.edit_message_text("\n".join(lines))
    return ConversationHandler.END


async def ui_delete_rule(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()

    keyboard: list[list[InlineKeyboardButton]] = []
    for i, cat in enumerate(UI_CATEGORIES):
        if i % 2 == 0:
            keyboard.append([])
        keyboard[-1].append(
            InlineKeyboardButton(cat["label"], callback_data=f"delcat:{cat['rule_key']}")
        )
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])
    await query.edit_message_text(
        "Выбери категорию, из которой нужно удалить ключевое слово.",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_DELETE_CATEGORY


async def ui_choose_delete_category(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()
    data = query.data or ""
    rule_key = data.split(":", 1)[1] if ":" in data else ""
    if not rule_key:
        await query.edit_message_text("Не удалось выбрать категорию. Попробуй ещё раз /start.")
        return ConversationHandler.END

    rules = load_rules()
    rule = rules.get(rule_key)
    if not rule:
        await query.edit_message_text("Категория не найдена.")
        return ConversationHandler.END

    context.user_data["delete_rule_key"] = rule_key
    keyboard = [[InlineKeyboardButton(kw, callback_data=f"delkw:{kw}")] for kw in rule.keywords]
    keyboard.append([InlineKeyboardButton("Отмена", callback_data="cancel")])
    await query.edit_message_text(
        "Выбери ключевое слово для удаления:",
        reply_markup=InlineKeyboardMarkup(keyboard),
    )
    return STATE_PICK_DELETE_KEYWORD


async def ui_delete_keyword(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    if not query:
        return ConversationHandler.END
    await query.answer()
    data = query.data or ""
    keyword = data.split(":", 1)[1] if ":" in data else ""
    rule_key = context.user_data.get("delete_rule_key")
    if not rule_key or not keyword:
        await query.edit_message_text("Не удалось удалить ключевое слово. Повтори через /start.")
        return ConversationHandler.END
    try:
        remove_keyword(rule_key, keyword)
    except Exception as e:
        await query.edit_message_text(f"Ошибка удаления: {e}")
        return ConversationHandler.END

    label = next((c["label"] for c in UI_CATEGORIES if c["rule_key"] == rule_key), rule_key)
    await query.edit_message_text(f"✅ Удалено из '{label}': {keyword}")
    context.user_data.pop("delete_rule_key", None)
    return ConversationHandler.END


def extract_date_from_filename(filename: str) -> str | None:
    """Извлекает дату из названия файла. Пример: 'Экспорт чеков от 17-01-2026.xlsx' -> '17.01.2026'"""
    # Паттерны: 17-01-2026, 17.01.2026, 2026-01-17
    for pattern in [
        r'(\d{2})[\-\.](\d{2})[\-\.](\d{4})',  # 17-01-2026 или 17.01.2026
        r'(\d{4})[\-\.](\d{2})[\-\.](\d{2})',  # 2026-01-17
    ]:
        m = re.search(pattern, filename)
        if m:
            g = m.groups()
            if len(g[0]) == 4:  # год первым
                return f"{g[2]}.{g[1]}.{g[0]}"
            return f"{g[0]}.{g[1]}.{g[2]}"
    return None


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Обработчик загрузки документа (Excel файла)"""
    document = update.message.document
    filename = document.file_name or ""

    if not (filename.endswith('.xlsx') or filename.endswith('.xls')):
        await update.message.reply_text(
            "⚠️ Пожалуйста, отправьте файл Excel (.xlsx или .xls)."
        )
        return

    await update.message.reply_text("⏳ Обрабатываю файл...")

    try:
        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
            await file.download_to_drive(tmp.name)
            tmp_path = tmp.name

        try:
            data = parse_excel_report(tmp_path)
            report_date = extract_date_from_filename(filename)
            if not report_date:
                report_date = datetime.now().strftime('%d.%m.%Y')
            report = format_report(data, report_date)
            await update.message.reply_text(f"✅ Готово!\n\n{report}")
        finally:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)

    except Exception as e:
        await update.message.reply_text(
            f"❌ Ошибка при обработке файла:\n{str(e)}\n\n"
            "Убедитесь, что файл — это экспорт чеков с корректной структурой "
            "(столбцы I — тип операции, O — список позиций)."
        )
        raise


def main() -> None:
    """Запуск бота"""
    # Создаём event loop для Python 3.10+ (иначе RuntimeError на MainThread)
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

    if BOT_TOKEN == 'YOUR_BOT_TOKEN_HERE':
        print("=" * 50)
        print("ОШИБКА: не задан токен бота.")
        print("На хостинге: Environment Variables / Secrets, добавьте одну из:")
        print("  REPORT_BOT_TOKEN=<токен от @BotFather>")
        print("  или BOT_TOKEN=<токен>")
        print("  или TELEGRAM_BOT_TOKEN=<токен>")
        print("Локально: скопируйте config.example.py -> config.py и укажите BOT_TOKEN.")
        print("=" * 50)
        sys.exit(1)

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    conv = ConversationHandler(
        entry_points=[
            CallbackQueryHandler(ui_add_rule, pattern=r"^ui:add_rule$"),
            CallbackQueryHandler(ui_delete_rule, pattern=r"^ui:delete_rule$"),
            CallbackQueryHandler(ui_show_rules, pattern=r"^ui:show_rules$"),
        ],
        states={
            STATE_PICK_CATEGORY: [CallbackQueryHandler(ui_choose_category, pattern=r"^cat:")],
            STATE_WAIT_KEYWORD: [MessageHandler(filters.TEXT & ~filters.COMMAND, ui_receive_keyword)],
            STATE_PICK_DELETE_CATEGORY: [
                CallbackQueryHandler(ui_choose_delete_category, pattern=r"^delcat:")
            ],
            STATE_PICK_DELETE_KEYWORD: [
                CallbackQueryHandler(ui_delete_keyword, pattern=r"^delkw:")
            ],
        },
        fallbacks=[
            CallbackQueryHandler(ui_cancel, pattern=r"^cancel$"),
            CommandHandler("start", start),
        ],
        allow_reentry=True,
    )
    app.add_handler(conv)

    try:
        rules = load_rules()
        print(f"BUILD: {APP_BUILD}")
        print("RULES: action_happy_hours =", rules["action_happy_hours"].keywords)
        print("RULES: action_last_hour  =", rules["action_last_hour"].keywords)
        print("RULES: advance_dr        =", rules["advance_dr"].keywords)
    except Exception as e:
        print("WARNING: failed to load rules at startup:", e)

    print("Бот запущен. Отправьте файл Excel для формирования отчёта.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
