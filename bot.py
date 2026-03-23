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

from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

from report_parser import parse_excel_report, format_report


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
    await update.message.reply_text(
        "Привет! Я бот для формирования вечернего отчёта.\n\n"
        "📤 Отправь мне файл Excel (.xlsx) с экспортом чеков — я сформирую отчёт и отправлю его тебе.\n\n"
        "Дата в отчёте будет взята из названия файла (например, 'Экспорт чеков от 17-01-2026.xlsx' → 17.01.2026), "
        "или сегодняшняя, если дату не удастся определить."
    )


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

    print("Бот запущен. Отправьте файл Excel для формирования отчёта.")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
