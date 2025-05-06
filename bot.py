import csv
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes
from telegram.ext import filters

TOKEN = "7919057808:AAG5bGV6fetfBKpM7khUaMqv9sMnY6oFy2w"  # Ganti dengan token bot Anda

# Fungsi untuk membaca CSV dan mengembalikan data dalam bentuk dictionary
def load_data_from_csv(file_name="data.csv"):
    data_dict = {}
    with open(file_name, mode='r', encoding='utf-8') as file:
        reader = csv.DictReader(file)
        for row in reader:
            id_bersih = row['NDEM'].strip().replace("'", "") 
            data_dict[id_bersih] = {"nama": row['CUSTOMER_NAME'], "NDEM": row['NDEM']}
    return data_dict

# Memuat data dari CSV
data_dict = load_data_from_csv()

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Halo! Kirim ID untuk melihat data (contoh: 1001).")

async def handle_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_input = update.message.text.strip()
    if user_input in data_dict:
        data = data_dict[user_input]
        response = f"🔍 NDEM {user_input}:\nNama: {data['nama']}\nEmail: {data['NDEM']}"
    else:
        response = "❌ ID tidak ditemukan."
    await update.message.reply_text(response)

def main():
    # Inisialisasi Application (menggantikan Updater)
    application = Application.builder().token(TOKEN).build()

    # Tambahkan handler
    application.add_handler(CommandHandler("start", start))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_id))

    # Jalankan bot
    application.run_polling()

if __name__ == "__main__":
    main()
