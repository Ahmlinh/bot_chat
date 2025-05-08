import pandas as pd
import requests
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from datetime import datetime
import pytz
import re
from flask import Flask
import threading

app = Flask('')


@app.route('/')
def home():
    return "Bot is running!"


def run():
    app.run(host='0.0.0.0', port=8080)


def keep_alive():
    t = threading.Thread(target=run)
    t.start()


keep_alive()

# Konfigurasi
SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxJAZsR6Mu0sLLFlOnDAzw_uDpGTckA7EjOW4z7rD_glRaUDMFMhkAfa9drKgGn0a0J/exec'
BOT_TOKEN = '7737796707:AAEy6NFX02PA2H1vH2cdptjNnZrU1uk6yQ0'
TIMEZONE = pytz.timezone('Asia/Jakarta')


def clean_ndem(value):
    """Menghilangkan tanda ' dari kolom NDEM"""
    if isinstance(value, str):
        return value.replace("'", "")
    return value


def load_data_from_apps_script():
    """Mengambil data dari Google Apps Script dan memprosesnya"""
    try:
        response = requests.get(SCRIPT_URL)
        response.raise_for_status()
        data = response.json()

        df = pd.DataFrame(data)

        # Membersihkan kolom NDEM
        if 'NDEM' in df.columns:
            df['NDEM'] = df['NDEM'].apply(clean_ndem)

        # Konversi kolom tanggal
        if 'TANGGAL PROSES' in df.columns:
            # Coba parsing dengan berbagai format
            df['TANGGAL PROSES'] = pd.to_datetime(
                df['TANGGAL PROSES'],
                format='mixed',  # Menerima berbagai format
                utc=True,  # Untuk timestamp dengan 'Z'
                dayfirst=True  # Prioritaskan format hari-pertama
            )

            # Konversi ke timezone lokal jika diperlukan
            if df['TANGGAL PROSES'].dt.tz is not None:
                df['TANGGAL PROSES'] = df['TANGGAL PROSES'].dt.tz_convert(
                    TIMEZONE)
            else:
                df['TANGGAL PROSES'] = df['TANGGAL PROSES'].dt.tz_localize(
                    TIMEZONE)

        return df

    except Exception as e:
        print(f"Error loading data: {str(e)}")
        return pd.DataFrame()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /start"""
    help_text = """
Selamat datang di Bot Analisis Transaksi!

Perintah yang tersedia:
/start - Menampilkan pesan bantuan
/total - Jumlah total transaksi
/per_cabang - Transaksi per cabang
/per_sales - Transaksi per sales
/per_sales_bulan - Transaksi per sales
/bulan YYYY-MM - Transaksi per bulan
/cari [nomor] - Cari data berdasarkan nomor
/exwitelmadiun - Laporan ex-Witel Madiun"""
    await update.message.reply_text(help_text)


async def cari_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk mencari data berdasarkan nomor"""
    if not context.args:
        await update.message.reply_text("Format salah. Gunakan: /cari [nomor]")
        return

    nomor = ' '.join(context.args)
    df = load_data_from_apps_script()

    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    # Cari di kolom NDEM yang sudah dibersihkan
    result = df[df['NDEM'].str.contains(nomor, case=False, na=False)]

    if result.empty:
        await update.message.reply_text(
            f"Tidak ditemukan data dengan nomor {nomor}")
        return

    # Format hasil pencarian
    response_text = f"🔍 Hasil pencarian untuk '{nomor}':\n\n"
    for _, row in result.head(5).iterrows():  # Batasi 5 hasil teratas
        response_text += (
            f"📅 Tanggal: {row['TANGGAL PROSES'].strftime('%d/%m/%Y %H:%M')}\n"
            f"🔢 NDEM: {row['NDEM']}\n"
            f"🏢 Cabang: {row.get('DATEL', 'N/A')}\n"
            f"👤 Sales: {row.get('NAMA SALES', 'N/A')}\n"
            f"👤 Nama Pelanggan: {row.get('CUSTOMER NAME', 'N/A')}\n")

    if len(result) > 5:
        response_text += f"\n⚠️ Menampilkan 5 dari {len(result)} hasil. Gunakan pencarian lebih spesifik."

    await update.message.reply_text(response_text)


async def total(update: Update, context: ContextTypes.DEFAULT_TYPE):
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    total_transaksi = len(df)
    await update.message.reply_text(
        f"Jumlah total transaksi: {total_transaksi:,}")


async def per_cabang(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /per_cabang"""
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    if 'DATEL' not in df.columns:
        await update.message.reply_text(
            "Kolom 'DATEL' tidak ditemukan dalam data.")
        return

    result = df['DATEL'].value_counts().sort_values(ascending=False)
    text = "📊 Jumlah transaksi per cabang:\n\n" + "\n".join(
        f"• {cabang}: {jumlah:,}" for cabang, jumlah in result.items())
    await update.message.reply_text(text)


async def per_sales(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /per_sales"""
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    if 'NAMA SALES' not in df.columns:
        await update.message.reply_text(
            "Kolom 'NAMA SALES' tidak ditemukan dalam data.")
        return

    result = df['NAMA SALES'].value_counts().sort_values(ascending=False)
    text = "👤 Jumlah transaksi per sales:\n\n" + "\n".join(
        f"• {sales}: {jumlah:,}" for sales, jumlah in result.items())
    await update.message.reply_text(text)


async def per_sales_bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /per_sales_bulan"""
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    if len(context.args) < 1:
        await update.message.reply_text(
            "Gunakan format: /per_sales_bulan [NAMA BULAN] (contoh: /per_sales_bulan November)"
        )
        return

    bulan_input = ' '.join(context.args).strip().lower()

    if 'BULAN' not in df.columns:
        await update.message.reply_text(
            "Kolom 'BULAN' tidak tersedia dalam data.")
        return

    df['BULAN'] = df['BULAN'].astype(str).str.lower()
    df_filtered = df[df['BULAN'].str.lower() == bulan_input]

    if df_filtered.empty:
        await update.message.reply_text(
            f"Tidak ditemukan data untuk bulan '{bulan_input.title()}'.")
        return

    if 'NAMA SALES' not in df_filtered.columns:
        await update.message.reply_text(
            "Kolom 'NAMA SALES' tidak tersedia dalam data.")
        return

    result = df_filtered['NAMA SALES'].value_counts().sort_values(
        ascending=False)
    text = f"📊 Jumlah transaksi per sales untuk bulan {bulan_input.title()}:\n\n" + "\n".join(
        f"• {sales}: {jumlah:,}" for sales, jumlah in result.items())
    await update.message.reply_text(text)


async def bulan(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /bulan"""
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    if len(context.args) != 1:
        await update.message.reply_text(
            "Format salah. Gunakan: /bulan YYYY-MM (contoh: /bulan 2023-11)")
        return

    try:
        bulan_input = datetime.strptime(context.args[0], "%Y-%m")
        bulan_str = bulan_input.strftime("%Y-%m")

        if 'TANGGAL PROSES' not in df.columns:
            await update.message.reply_text(
                "Kolom tanggal tidak ditemukan dalam data.")
            return

        df_bulan = df[df['TANGGAL PROSES'].dt.strftime('%Y-%m') == bulan_str]

        if df_bulan.empty:
            await update.message.reply_text(
                f"Tidak ada transaksi untuk bulan {bulan_str}")
            return

        total_transaksi = len(df_bulan)
        text = f"📅 Transaksi bulan {bulan_str}:\n\n" \
               f"• Jumlah transaksi: {total_transaksi:,}\n" \
               f"• Per cabang:\n"

        per_cabang = df_bulan['DATEL'].value_counts().sort_values(
            ascending=False)
        text += "\n".join(f"  - {cabang}: {jumlah:,}"
                          for cabang, jumlah in per_cabang.items())

        await update.message.reply_text(text)

    except ValueError:
        await update.message.reply_text(
            "Format tanggal salah. Gunakan: /bulan YYYY-MM (contoh: /bulan 2023-11)"
        )
    except Exception as e:
        await update.message.reply_text(f"Terjadi error: {str(e)}")


async def ex_witel_madiun(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handler untuk command /ex-witel_madiun"""
    df = load_data_from_apps_script()
    if df.empty:
        await update.message.reply_text(
            "Gagal memuat data. Silakan coba lagi nanti.")
        return

    # Datel yang ditampilkan
    target_datel = ["MADIUN", "BOJONEGORO", "TUBAN", "NGAWI", "PONOROGO"]

    # Validasi kolom penting
    for col in ['DATEL', 'BULAN', 'ADDON']:
        if col not in df.columns:
            await update.message.reply_text(
                f"Kolom '{col}' tidak ditemukan dalam data.")
            return

    # Standardisasi data
    df['DATEL'] = df['DATEL'].astype(str).str.upper()
    df['BULAN'] = df['BULAN'].astype(str).str.title()
    df['ADDON'] = df['ADDON'].astype(str).str.title()

    # Filter hanya DATEL tertentu
    df_filtered = df[df['DATEL'].isin(target_datel)]

    if df_filtered.empty:
        await update.message.reply_text(
            "Tidak ada data untuk DATEL yang diminta.")
        return

    response = "📊 **Laporan Witel Madiun**\n\n"
    bulan_urut = [
        'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
        'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ]

    for datel in target_datel:
        sub_df = df_filtered[df_filtered['DATEL'] == datel]
        response += f"🏢\n\n**{datel.title()}**\n"
        response += f"• Total pelanggan: {len(sub_df):,}\n"
        response += f"• Pelanggan per bulan:\n\n"

        # Buat grouped data: BULAN, ADDON, JUMLAH
        grouped = sub_df.groupby(['BULAN',
                                  'ADDON']).size().reset_index(name='JUMLAH')

        # Loop berdasarkan urutan bulan
        for bulan in bulan_urut:
            bulan_df = sub_df[sub_df['BULAN'] == bulan]
            if not bulan_df.empty:
                total_bulan = len(bulan_df)
                response += f"* {bulan}: {total_bulan:,}\n"

                # Rincian ADDON per bulan
                addon_rows = grouped[grouped['BULAN'] == bulan]
                for _, row in addon_rows.iterrows():
                    response += f"  • {row['ADDON']}: {row['JUMLAH']:,}\n"

        response += "\n"  # Spasi antar datel

    await update.message.reply_text(response[:4000]
                                    )  # Batasi 4000 karakter Telegram

    def run_bot():
        """Menjalankan bot Telegram dalam thread terpisah"""
        app_bot = ApplicationBuilder().token(BOT_TOKEN).build()

        # Daftar handler command
        app_bot.add_handler(CommandHandler("start", start))
        app_bot.add_handler(CommandHandler("total", total))
        app_bot.add_handler(CommandHandler("per_cabang", per_cabang))
        app_bot.add_handler(CommandHandler("per_sales", per_sales))
        app_bot.add_handler(CommandHandler("per_sales_bulan", per_sales_bulan))
        app_bot.add_handler(CommandHandler("bulan", bulan))
        app_bot.add_handler(CommandHandler("cari", cari_data))
        app_bot.add_handler(CommandHandler("exwitelmadiun", ex_witel_madiun))

        print("🤖 Bot sedang berjalan...")
        app_bot.run_polling()

    def main():
        keep_alive()  # Jalankan Flask dulu
        threading.Thread(
            target=run_bot).start()  # Jalankan bot di thread terpisah

    if __name__ == '__main__':
        main()
