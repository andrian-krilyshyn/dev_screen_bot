import fitz  # PyMuPDF
import io
import asyncio
from googleapiclient.discovery import build
from google.oauth2 import service_account
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from telegram import Bot, Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, CallbackContext, ContextTypes, CallbackQueryHandler
from datetime import datetime

# Ваш токен бота та ID чату
TOKEN = '7320369617:AAGitENCsLTIjqL9bPU-P7gJyC0fgSStz2U'
CHAT_ID = '-1002421170446'

# Налаштування для Google Sheets
SERVICE_ACCOUNT_FILE = 's.json'
SCOPES = ['https://www.googleapis.com/auth/drive']
SHEET_ID = '1_8vNN53eZMHwQ9w3mZySs0560cbpl-kX8MoUc94eNR8'

# Налаштування Google Drive API
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('drive', 'v3', credentials=credentials)

async def fetch_pdf():
    request = service.files().export_media(fileId=SHEET_ID, mimeType='application/pdf')
    fh = io.BytesIO(request.execute())
    with open('output.pdf', 'wb') as f:
        f.write(fh.getvalue())

def crop_pdf(input_pdf_path, output_pdf_path, crop_area, zoom):
    doc = fitz.open(input_pdf_path)
    page = doc.load_page(0)  # Завантажуємо першу сторінку
    
    # Якщо передано crop_area, використовуємо його, інакше беремо стандартний прямокутник
    if crop_area is None:
        crop_area = fitz.Rect(50, 50, page.rect.width - 50, page.rect.height - 50)

    # Застосовуємо масштаб для підвищення якості
    pix = page.get_pixmap(clip=crop_area, matrix=fitz.Matrix(zoom, zoom))

    # Створюємо новий PDF і додаємо обрізану сторінку
    cropped_doc = fitz.open()
    new_page = cropped_doc.new_page(width=pix.width, height=pix.height)
    new_page.insert_image(fitz.Rect(0, 0, pix.width, pix.height), stream=pix.tobytes())
    
    # Зберігаємо обрізаний PDF
    cropped_doc.save(output_pdf_path)
    cropped_doc.close()

async def send_pdf(pdf_path):
    bot = Bot(token=TOKEN)
    with open(pdf_path, 'rb') as pdf_file:
        sent_message = await bot.send_document(chat_id=CHAT_ID, document=pdf_file)
        table_name = "CO PEDRO DANIL"  # Змініть на відповідне джерело назви таблиці, якщо потрібно
    # Час надсилання
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    # Надсилаємо текстове повідомлення як відповідь на надісланий PDF
    if sent_message:
        await bot.send_message(chat_id=CHAT_ID, text=f"Документ '{table_name}' был послан в {current_time}.", reply_to_message_id=sent_message.message_id)
    else:
        await bot.send_message(chat_id=CHAT_ID, text=f"Документ '{table_name}' был послан в {current_time}.")

async def job():
    await fetch_pdf()

    doc = fitz.open('output.pdf')
    page = doc.load_page(0)
    
    # Визначаємо область обрізки як об'єкт fitz.Rect
    crop_area = fitz.Rect(50, 50, 0.30 * page.rect.width, page.rect.height*0.55)
    
    # Використовуємо crop_area з правильним форматом
    crop_pdf('output.pdf', 'CO_PEDRO_DANIL.pdf', crop_area, zoom=6)

    await send_pdf('CO_PEDRO_DANIL.pdf')

async def start(update: Update, context: CallbackContext):
    keyboard = [[InlineKeyboardButton("Выполнить действия", callback_data='run_job')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text('Нажмите кнопку для выполнения действий.', reply_markup=reply_markup)

async def button(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()
    if query.data == 'run_job':
        await job()

def main():
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(CommandHandler('start', start))
    app.add_handler(CallbackQueryHandler(button))

    # Запуск асинхронного циклу планувальника
    scheduler = AsyncIOScheduler()
    scheduler.add_job(job, 'interval', minutes=30)
    scheduler.start()

    app.run_polling()

if __name__ == '__main__':
    main()
