import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
    CallbackQueryHandler,
    ConversationHandler,
)
import io
import logging
from docx2pdf import convert
import pythoncom
import tempfile
import os
from PIL import Image
import pytesseract
import requests
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
import re

# Setup logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

SELECT_PAGES, CONFIRM_PAGES = range(2)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /start is issued."""
    await update.message.reply_text(
        'Halo! Saya adalah bot konversi dokumen. Saya dapat:\n\n'
        '1. Mengkonversi file Excel/CSV/TXT ke format Excel standar (.xlsx)\n'
        '2. Mengkonversi file Word (.docx) ke PDF\n'
        '3. Mengekstrak teks dari gambar (OCR)\n'
        '4. Mengkonversi gambar (JPG/PNG/WEBP) ke PDF\n'
        '5. Mengekstrak halaman tertentu dari PDF\n\n'
        'Gunakan perintah:\n'
        '/excel - Konversi ke Excel\n'
        '/word2pdf - Konversi Word ke PDF\n'
        '/ocr - Ekstrak teks dari gambar\n'
        '/img2pdf - Konversi gambar ke PDF\n'
        '/extractpdf - Ekstrak halaman dari PDF\n\n'
        'Atau langsung unggah file yang ingin dikonversi.'
    )



async def excel_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Explain how to convert to Excel."""
    await update.message.reply_text(
        'Kirimkan saya file Excel (.xls, .xlsx), CSV (.csv), atau file teks tab-delimited (.txt) '
        'dan saya akan mengkonversinya ke format Excel (.xlsx) standar.'
    )

async def word2pdf_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Explain how to convert Word to PDF."""
    await update.message.reply_text(
        'Kirimkan saya file Word (.docx) dan saya akan mengkonversinya ke format PDF.'
    )


async def ocr_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    context.user_data['mode'] = 'ocr'  # ✅ Tambahkan baris ini
    await update.message.reply_text(
        'Kirimkan saya gambar (format JPG/PNG) dan saya akan mengekstrak teks yang ada di dalamnya.\n\n'
        'Anda juga bisa mengirim beberapa gambar sekaligus, dan saya akan menggabungkan semua teks yang ditemukan.'
    )

async def extract_pdf_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Explain how to extract pages from PDF."""
    await update.message.reply_text(
        'Kirimkan saya file PDF dan saya akan membantu Anda mengekstrak halaman tertentu.\n\n'
        'Setelah mengirim PDF, saya akan menunjukkan jumlah halaman dan meminta Anda memilih halaman yang ingin diekstrak.'
    )
    return SELECT_PAGES

async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle PDF file upload for extraction."""
    file = update.message.document
    
    if not file.file_name.lower().endswith('.pdf'):
        await update.message.reply_text('File harus dalam format PDF.')
        return ConversationHandler.END
    
    # Download the PDF file
    file_id = file.file_id
    new_file = await context.bot.get_file(file_id)
    file_bytes = await new_file.download_as_bytearray()
    
    # Read the PDF to get page count
    pdf_reader = PdfReader(io.BytesIO(file_bytes))
    page_count = len(pdf_reader.pages)
    
    # Store PDF data in user context
    context.user_data['pdf_data'] = file_bytes
    context.user_data['page_count'] = page_count
    context.user_data['original_name'] = file.file_name
    
    await update.message.reply_text(
        f'PDF yang Anda upload memiliki {page_count} halaman.\n\n'
        'Silakan masukkan nomor halaman yang ingin Anda ekstrak, contoh:\n'
        '- "1" untuk halaman 1 saja\n'
        '- "1,3,5" untuk beberapa halaman tertentu\n'
        '- "1-5" untuk range halaman\n'
        '- "1,3-5,7" untuk kombinasi\n\n'
        'Balas dengan "batal" untuk membatalkan.'
    )
    
    return CONFIRM_PAGES

async def confirm_pages(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Process page selection and extract pages."""
    user_input = update.message.text.strip().lower()
    
    if user_input == 'batal':
        await update.message.reply_text('Ekstraksi PDF dibatalkan.')
        return ConversationHandler.END
    
    try:
        # Parse page selection
        page_count = context.user_data['page_count']
        selected_pages = parse_page_selection(user_input, page_count)
        
        if not selected_pages:
            await update.message.reply_text(
                'Format pemilihan halaman tidak valid. Silakan coba lagi atau balas "batal" untuk membatalkan.'
            )
            return CONFIRM_PAGES
        
        # Extract pages
        pdf_data = context.user_data['pdf_data']
        pdf_reader = PdfReader(io.BytesIO(pdf_data))
        pdf_writer = PdfWriter()
        
        for page_num in selected_pages:
            pdf_writer.add_page(pdf_reader.pages[page_num - 1])
        
        # Prepare output
        output = io.BytesIO()
        pdf_writer.write(output)
        output.seek(0)
        
        original_name = context.user_data['original_name']
        base_name = original_name[:-4] if original_name.lower().endswith('.pdf') else original_name
        output_name = f"{base_name}_halaman_{user_input.replace(',', '_').replace('-', '_')}.pdf"
        
        # Send extracted PDF
        await update.message.reply_document(
            document=output,
            filename=output_name,
            caption=f'Berhasil mengekstrak halaman: {user_input}'
        )
        
        # Clean up
        context.user_data.pop('pdf_data', None)
        context.user_data.pop('page_count', None)
        context.user_data.pop('original_name', None)
        
    except Exception as e:
        logger.error(f"Error extracting PDF pages: {e}")
        await update.message.reply_text(
            'Terjadi kesalahan saat mengekstrak halaman. Silakan coba lagi atau balas "batal" untuk membatalkan.'
        )
        return CONFIRM_PAGES
    
    return ConversationHandler.END

def parse_page_selection(input_str: str, max_pages: int) -> list:
    """Parse user input for page selection."""
    if not input_str:
        return []
    
    # Remove all whitespace
    input_str = input_str.replace(' ', '')
    
    # Split by commas
    parts = input_str.split(',')
    pages = set()
    
    for part in parts:
        if '-' in part:
            # Handle range (e.g., 1-5)
            range_parts = part.split('-')
            if len(range_parts) != 2:
                return []
            
            try:
                start = int(range_parts[0])
                end = int(range_parts[1])
            except ValueError:
                return []
            
            if start < 1 or end > max_pages or start > end:
                return []
            
            pages.update(range(start, end + 1))
        else:
            # Handle single page
            try:
                page = int(part)
            except ValueError:
                return []
            
            if page < 1 or page > max_pages:
                return []
            
            pages.add(page)
    
    return sorted(pages)

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel the conversation."""
    await update.message.reply_text('Operasi dibatalkan.')
    
    # Clean up
    context.user_data.pop('pdf_data', None)
    context.user_data.pop('page_count', None)
    context.user_data.pop('original_name', None)
    
    return ConversationHandler.END

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = update.message.document
    user_mode = context.user_data.get("mode", None)

    if file:
        file_name = file.file_name.lower()
        if file_name.endswith(('.xls', '.xlsx', '.csv', '.txt')):
            await handle_excel_conversion(update, context)
        elif file_name.endswith(('.docx', '.doc')):
            await handle_word_to_pdf(update, context)
        elif file_name.endswith(('.jpg', '.jpeg', '.png', '.webp', '.bmp')):
            # Jalankan langsung jika mode ditentukan
            if user_mode == 'img2pdf':
                await handle_image_to_pdf(update, context)
            elif user_mode == 'ocr':
                await handle_image_ocr(update, context)
            else:
                await update.message.reply_text(
                    'Anda mengirim gambar. Apa yang ingin Anda lakukan?\n\n'
                    '1. Ekstrak teks dari gambar - gunakan /ocr\n'
                    '2. Konversi gambar ke PDF - gunakan /img2pdf\n\n'
                    'Atau balas dengan perintah yang diinginkan.'
                )
        else:
            await update.message.reply_text('Format file tidak didukung.')
    
    elif update.message.photo:
        mode = context.user_data.get('mode')
        if mode == 'ocr':
            await handle_image_ocr(update, context)
            context.user_data.pop('mode', None)
        elif mode == 'img2pdf':
            await handle_image_to_pdf(update, context)
            context.user_data.pop('mode', None)
        else:
            await update.message.reply_text(
                'Anda mengirim gambar. Apa yang ingin Anda lakukan?\n\n'
                '1. Ekstrak teks dari gambar - gunakan /ocr\n'
                '2. Konversi gambar ke PDF - gunakan /img2pdf\n\n'
                'Atau balas dengan perintah yang diinginkan.'
            )

# Command /img2pdf
async def image_to_pdf_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [
            InlineKeyboardButton("🔵 Portrait", callback_data='pdf_orientation:portrait'),
            InlineKeyboardButton("🔹 Landscape", callback_data='pdf_orientation:landscape')
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await update.message.reply_text(
        "Silakan pilih orientasi halaman untuk PDF:",
        reply_markup=reply_markup
    )

# Callback pilihan orientasi
async def handle_pdf_orientation_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    orientation = query.data.split(":")[1]
    context.user_data["pdf_orientation"] = orientation

    await query.edit_message_text(
        f"Orientasi PDF disetel ke: {orientation.capitalize()}. Sekarang kirim gambar yang ingin dikonversi."
    )


async def handle_image_to_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if not update.message.photo and not update.message.document:
            await update.message.reply_text("Tidak ada gambar terdeteksi. Kirim gambar yang ingin dikonversi.")
            return

        orientation = context.user_data.get("pdf_orientation", "portrait")

        if update.message.document:
            file = update.message.document
            file_id = file.file_id
            new_file = await context.bot.get_file(file_id)
            image_bytes = await new_file.download_as_bytearray()
            images = [Image.open(io.BytesIO(image_bytes))]
        elif update.message.photo:
            photo = update.message.photo[-1]
            file_id = photo.file_id
            new_file = await context.bot.get_file(file_id)
            image_bytes = await new_file.download_as_bytearray()
            images = [Image.open(io.BytesIO(image_bytes))]
        else:
            await update.message.reply_text("Gagal membaca gambar.")
            return

        # Siapkan PDF
        pdf = FPDF(orientation='P' if orientation == 'portrait' else 'L', unit='mm', format='A4')
        margin = 10  # mm

        for img in images:
            # Dapatkan ukuran halaman PDF
            page_width = 210 if orientation == 'portrait' else 297
            page_height = 297 if orientation == 'portrait' else 210

            available_width = page_width - 2 * margin
            available_height = page_height - 2 * margin

            # Ukuran asli gambar (dalam pixel)
            img_width, img_height = img.size
            img_ratio = img_width / img_height
            page_ratio = available_width / available_height

            # Resize agar muat di halaman dengan margin, pertahankan rasio
            if img_ratio > page_ratio:
                new_width = available_width
                new_height = available_width / img_ratio
            else:
                new_height = available_height
                new_width = available_height * img_ratio

            # Simpan gambar sementara sebagai PNG
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_img:
                img.save(temp_img.name, "PNG")
                temp_img_path = temp_img.name

            # Tambahkan halaman dan gambar di tengah halaman
            pdf.add_page()
            x_offset = (page_width - new_width) / 2
            y_offset = (page_height - new_height) / 2
            pdf.image(temp_img_path, x_offset, y_offset, new_width, new_height)
            os.unlink(temp_img_path)

        pdf_bytes = io.BytesIO()
        pdf.output(pdf_bytes)
        pdf_bytes.seek(0)

        await update.message.reply_document(
            document=pdf_bytes,
            filename="converted_image.pdf",
            caption=f"Gambar berhasil dikonversi ke PDF dengan orientasi {orientation.capitalize()}"
        )

        # Reset orientasi untuk sesi berikutnya
        context.user_data.pop("pdf_orientation", None)

    except Exception as e:
        await update.message.reply_text("Terjadi kesalahan saat mengkonversi gambar ke PDF.")

async def handle_image_ocr(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle image OCR to extract and format paragraphs."""
    try:
        if not update.message.photo and not update.message.document:
            await update.message.reply_text("Tidak ada gambar terdeteksi. Silakan kirim gambar yang jelas.")
            return

        # Check if it's a document (file) or photo
        if update.message.document:
            file = update.message.document
            file_id = file.file_id
            new_file = await context.bot.get_file(file_id)
            image_bytes = await new_file.download_as_bytearray()
            image = Image.open(io.BytesIO(image_bytes))
        else:
            photo = update.message.photo[-1]
            file_id = photo.file_id
            new_file = await context.bot.get_file(file_id)
            image_bytes = await new_file.download_as_bytearray()
            image = Image.open(io.BytesIO(image_bytes))

        # Ambil data OCR dengan detail
        ocr_data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)

        # Gabungkan teks berdasarkan nomor paragraf
        n_boxes = len(ocr_data['level'])
        paragraphs = {}
        for i in range(n_boxes):
            par_num = ocr_data['par_num'][i]
            text = ocr_data['text'][i].strip()

            if text:
                if par_num not in paragraphs:
                    paragraphs[par_num] = []
                paragraphs[par_num].append(text)

        # Gabungkan semua paragraf menjadi satu teks
        extracted_text = "\n\n".join(" ".join(words) for words in paragraphs.values())

        if not extracted_text.strip():
            await update.message.reply_text(
                'Tidak ada teks yang terdeteksi dalam gambar. '
                'Pastikan gambar jelas dan memiliki teks yang cukup besar.'
            )
            return

        formatted_text = f"📄 Hasil Ekstraksi Teks:\n\n{extracted_text}"

        if len(formatted_text) < 4000:
            await update.message.reply_text(formatted_text)
        else:
            with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as temp_file:
                temp_file.write(formatted_text.encode('utf-8'))
                temp_file_path = temp_file.name

            with open(temp_file_path, 'rb') as file_to_send:
                await update.message.reply_document(
                    document=file_to_send,
                    filename="hasil_ocr.txt",
                    caption="Hasil teks terlalu panjang, dikirim sebagai file."
                )

            os.unlink(temp_file_path)
            context.user_data.pop('mode', None)


    except Exception as e:
        logger.error(f"Error in OCR processing: {e}")
        await update.message.reply_text(
            'Terjadi kesalahan saat memproses gambar. '
            'Pastikan gambar jelas dan format didukung (JPG/PNG).'
        )


async def handle_excel_conversion(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle Excel/CSV/TXT to Excel conversion."""
    file = update.message.document
    
    # Download the file
    file_id = file.file_id
    new_file = await context.bot.get_file(file_id)
    file_bytes = await new_file.download_as_bytearray()
    
    # Try to read the file
    df = None
    file_type = None
    
    # Try as Excel first
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        file_type = 'Excel'
    except Exception as e:
        logger.info(f"Failed to read as Excel: {e}")
        # Try as CSV
        try:
            df = pd.read_csv(io.BytesIO(file_bytes))
            file_type = 'CSV'
        except Exception as e:
            logger.info(f"Failed to read as CSV: {e}")
            # Try as tab-delimited
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), sep='\t')
                file_type = 'Tab-delimited'
            except Exception as e:
                logger.info(f"Failed to read as tab-delimited: {e}")
                await update.message.reply_text(
                    'Gagal membaca file. Format tidak dikenali. '
                    'Pastikan file dalam format Excel, CSV, atau tab-delimited.'
                )
                return
    
    # If we successfully read the file, convert to .xlsx
    if df is not None:
        try:
            # Prepare the output file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            output.seek(0)
            
            # Determine the output filename
            original_name = file.file_name
            if '.' in original_name:
                base_name = original_name.split('.')[0]
            else:
                base_name = original_name
            output_name = f"{base_name}_converted.xlsx"
            
            # Send the file back
            await update.message.reply_document(
                document=output,
                filename=output_name,
                caption=f'File berhasil dikonversi dari {file_type} ke .xlsx'
            )
        except Exception as e:
            logger.error(f"Error converting file: {e}")
            await update.message.reply_text(
                'Terjadi kesalahan saat mengkonversi file. Silakan coba lagi.'
            )

async def handle_word_to_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle Word to PDF conversion."""
    file = update.message.document
    
    # Download the file
    file_id = file.file_id
    new_file = await context.bot.get_file(file_id)
    
    # Create a temporary directory
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Download the Word file
            word_path = os.path.join(temp_dir, file.file_name)
            await new_file.download_to_drive(word_path)
            
            # Convert to PDF
            pdf_path = os.path.join(temp_dir, 'converted.pdf')
            
            # Initialize COM for docx2pdf
            pythoncom.CoInitialize()
            convert(word_path, pdf_path)
            pythoncom.CoUninitialize()
            
            # Read the PDF file
            with open(pdf_path, 'rb') as f:
                pdf_bytes = f.read()
            
            # Determine the output filename
            original_name = file.file_name
            if '.' in original_name:
                base_name = original_name.split('.')[0]
            else:
                base_name = original_name
            output_name = f"{base_name}_converted.pdf"
            
            # Send the PDF file back
            await update.message.reply_document(
                document=io.BytesIO(pdf_bytes),
                filename=output_name,
                caption='File Word berhasil dikonversi ke PDF'
            )
        except Exception as e:
            logger.error(f"Error converting Word to PDF: {e}")
            await update.message.reply_text(
                'Terjadi kesalahan saat mengkonversi Word ke PDF. '
                'Pastikan file adalah dokumen Word yang valid (.docx).'
            )

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Log errors."""
    logger.error(f'Update {update} caused error {context.error}')

def main() -> None:
    """Start the bot."""
    # Replace 'YOUR_TOKEN' with your actual bot token
    application = Application.builder().token("7761036046:AAEPGXNke-m56Xi4MIs9u6fKm3nDs7sYu6M").build()

    pdf_extract_conv = ConversationHandler(
        entry_points=[CommandHandler("extractpdf", extract_pdf_command)],
        states={
            SELECT_PAGES: [MessageHandler(filters.Document.PDF, handle_pdf)],
            CONFIRM_PAGES: [MessageHandler(filters.TEXT & ~filters.COMMAND, confirm_pages)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )    
    
    # Register handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("excel", excel_command))
    application.add_handler(CommandHandler("word2pdf", word2pdf_command))
    application.add_handler(CommandHandler("ocr", ocr_command))
    application.add_handler(pdf_extract_conv)
    application.add_handler(CallbackQueryHandler(handle_pdf_orientation_callback, pattern="^pdf_orientation:"))
    application.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, handle_image_to_pdf))
    application.add_handler(MessageHandler(filters.Document.ALL | filters.PHOTO, handle_document))
    application.add_handler(CommandHandler("img2pdf", image_to_pdf_start))
    application.add_error_handler(error_handler)
    
    # Start the bot
    application.run_polling()

if __name__ == '__main__':
    main()
