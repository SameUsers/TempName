import os
import time
import telebot
from classes import Pdf_Worker

TOKEN = "7235411692:AAFGufOh_Jmd5Z5MSpPGf7Cmhjdc6VOq4ho"
bot = telebot.TeleBot(TOKEN)

SAVE_FOLDER = "pdf_files"
os.makedirs(SAVE_FOLDER, exist_ok=True)
XLSX_FOLDER = "xlsx_files"
os.makedirs(XLSX_FOLDER, exist_ok=True)

worker = Pdf_Worker()

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Ожидаю файлы Invoice/Specification/PL.")

@bot.message_handler(content_types=['document'])
def handle_files(message):
    doc = message.document
    if not doc.file_name.lower().endswith(".pdf"):
        bot.reply_to(message, "Только PDF файлы.")
        return

    # Отправляем прогресс-сообщение
    progress_msg = bot.send_message(message.chat.id, f"Начало обработки {doc.file_name}...\n[          ] 0%")

    def update_progress(percent):
        bar_length = 10
        filled_length = int(bar_length * percent / 100)
        bar = '█' * filled_length + ' ' * (bar_length - filled_length)
        bot.edit_message_text(chat_id=message.chat.id, message_id=progress_msg.message_id,
                              text=f"Обработка {doc.file_name}...\n[{bar}] {percent}%")

    # Шаг 1: Скачивание файла
    update_progress(10)
    file_info = bot.get_file(doc.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    file_path = os.path.join(SAVE_FOLDER, doc.file_name)
    with open(file_path, "wb") as f:
        f.write(downloaded_file)
    update_progress(30)

    # Определяем тип файла
    filename_lower = doc.file_name.lower()
    filter_spec = filename_lower == "specification_sell.pdf"
    remove_edges = filename_lower == "pl.pdf"
    invoice_lines = filename_lower == "invoice_purchase.pdf"

    # Шаг 2: Конвертация в XLSX
    xlsx_path = os.path.join(XLSX_FOLDER, doc.file_name.replace(".pdf", ".xlsx"))
    update_progress(50)
    success = worker.pdf_to_xlsx(
        file_path,
        xlsx_path,
        filter_spec=filter_spec,
        remove_edges=remove_edges,
        invoice_lines=invoice_lines
    )
    update_progress(100)
    time.sleep(0.5)  # небольшая задержка, чтобы пользователь увидел 100%

    # Финальное сообщение
    if success:
        bot.edit_message_text(chat_id=message.chat.id, message_id=progress_msg.message_id,
                              text=f"{doc.file_name} успешно преобразован в XLSX:\n{xlsx_path}")
    else:
        bot.edit_message_text(chat_id=message.chat.id, message_id=progress_msg.message_id,
                              text=f"Не удалось извлечь таблицы из {doc.file_name}.")

bot.infinity_polling()