import os
import time
import telebot
from classes import PdfWorker, FileGenerate, DocxFiller

TOKEN = "7235411692:AAFGufOh_Jmd5Z5MSpPGf7Cmhjdc6VOq4ho"
bot = telebot.TeleBot(TOKEN)

SAVE_FOLDER = "pdf_files"
os.makedirs(SAVE_FOLDER, exist_ok=True)
XLSX_FOLDER = "xlsx_files"
os.makedirs(XLSX_FOLDER, exist_ok=True)

worker = PdfWorker()
file_gen = FileGenerate()
docx_filler = DocxFiller()


@bot.message_handler(commands=['start'])
def start(message):
    """Обработчик команды /start - приветственное сообщение"""
    bot.send_message(message.chat.id, "Ожидаю файлы Invoice/Specification/PL.")


@bot.message_handler(content_types=['document'])
def handle_files(message):
    """Обработчик входящих PDF-файлов с конвертацией в XLSX и генерацией отчетов"""
    doc = message.document
    if not doc.file_name.lower().endswith(".pdf"):
        bot.reply_to(message, "Только PDF файлы.")
        return

    progress_msg = bot.send_message(
        message.chat.id,
        f"Начало обработки {doc.file_name}...\n[          ] 0%"
    )

    def update_progress(percent: int):
        """Обновление индикатора прогресса обработки файла"""
        bar_length = 10
        filled_length = int(bar_length * percent / 100)
        bar = '█' * filled_length + ' ' * (bar_length - filled_length)
        bot.edit_message_text(
            chat_id=message.chat.id,
            message_id=progress_msg.message_id,
            text=f"Обработка {doc.file_name}...\n[{bar}] {percent}%"
        )

    try:
        update_progress(10)
        
        file_info = bot.get_file(doc.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        file_path = os.path.join(SAVE_FOLDER, doc.file_name)
        
        with open(file_path, "wb") as f:
            f.write(downloaded_file)
        update_progress(30)

        filename_lower = doc.file_name.lower()
        filter_spec = filename_lower == "specification_sell.pdf"
        remove_edges = filename_lower == "pl.pdf"
        invoice_lines = filename_lower == "invoice_purchase.pdf"

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
        time.sleep(0.5)

        if not success:
            bot.edit_message_text(
                chat_id=message.chat.id,
                message_id=progress_msg.message_id,
                text=f"Не удалось извлечь таблицы из {doc.file_name}"
            )
            return

        bot.edit_message_text(
            chat_id=message.chat.id,
            message_id=progress_msg.message_id,
            text=f"{doc.file_name} успешно обработан. Генерация отчетов..."
        )

        output_filename = "123456789_invoice_sell_filled.xlsx"
        file_gen.fill_invoice(
            template_filename="123456789_invoice_sell.xlsx",
            invoice_filename=os.path.basename(xlsx_path),
            ref_filename="Справочник.xlsx",
            pl_filename="PL.xlsx",
            spec_filename="Specification_sell.xlsx",
            output_filename=output_filename
        )
        
        output_path = os.path.join("examples", output_filename)
        word_template = "examples/Description.docx"
        word_output = "examples/Description_Result.docx"

        success_word = docx_filler.fill_table_from_excel(
            template_path=word_template,
            excel_path=output_path,
            output_path=word_output,
            table_index=0
        )

        if success_word:
            with open(output_path, "rb") as f:
                bot.send_document(message.chat.id, f, caption="Готовый XLSX")
            
            with open(word_output, "rb") as f:
                bot.send_document(message.chat.id, f, caption="Готовый DOCX")
        else:
            bot.send_message(message.chat.id, "Ошибка при генерации Word документа")

    except Exception as e:
        error_msg = f"Ошибка обработки файла {doc.file_name}: {str(e)}"
        bot.edit_message_text(
            chat_id=message.chat.id,
            message_id=progress_msg.message_id,
            text=error_msg
        )


bot.infinity_polling(timeout=3600)