import sqlite3
import os
from datetime import datetime
from fpdf import FPDF
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

DB_NAME = "kas_bot.db"

# =============================================
# DATABASE
# =============================================

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT NOT NULL,
            jenis TEXT NOT NULL,
            nominal REAL NOT NULL,
            keterangan TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def simpan_transaksi(jenis: str, nominal: float, keterangan: str):
    tanggal = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO transaksi (tanggal, jenis, nominal, keterangan) VALUES (?, ?, ?, ?)",
        (tanggal, jenis, nominal, keterangan),
    )
    conn.commit()
    conn.close()

def hitung_saldo():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT COALESCE(SUM(nominal), 0) FROM transaksi WHERE jenis = 'pemasukan'")
    total_pemasukan = cursor.fetchone()[0]
    cursor.execute("SELECT COALESCE(SUM(nominal), 0) FROM transaksi WHERE jenis = 'pengeluaran'")
    total_pengeluaran = cursor.fetchone()[0]
    conn.close()
    return total_pemasukan - total_pengeluaran

def get_riwayat(limit=10):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT tanggal, jenis, nominal, keterangan
        FROM transaksi
        ORDER BY created_at DESC
        LIMIT ?
    """, (limit,))
    result = cursor.fetchall()
    conn.close()
    return result

def generate_pdf(transaksi, periode="Semua Transaksi"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Laporan Keuangan", 0, 1, "C")
    pdf.cell(0, 10, periode, 0, 1, "C")
    pdf.ln(10)

    pdf.set_font("Arial", "B", 12)
    pdf.cell(30, 10, "Tanggal", 1)
    pdf.cell(20, 10, "Jenis", 1)
    pdf.cell(30, 10, "Nominal", 1)
    pdf.cell(110, 10, "Keterangan", 1)
    pdf.ln()

    pdf.set_font("Arial", "", 10)
    total_pemasukan = 0
    total_pengeluaran = 0

    for tgl, jenis, nominal, ket in transaksi:
        pdf.cell(30, 10, tgl, 1)
        pdf.cell(20, 10, jenis, 1)
        symbol = "+" if jenis == "pemasukan" else "-"
        pdf.cell(30, 10, f"{symbol}{nominal:,.0f}", 1)
        pdf.cell(110, 10, ket or "-", 1)
        pdf.ln()

        if jenis == "pemasukan":
            total_pemasukan += nominal
        else:
            total_pengeluaran += nominal

    saldo = total_pemasukan - total_pengeluaran
    pdf.ln(10)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Total Pemasukan: Rp {total_pemasukan:,.0f}", 0, 1)
    pdf.cell(0, 10, f"Total Pengeluaran: Rp {total_pengeluaran:,.0f}", 0, 1)
    pdf.cell(0, 10, f"Saldo Bersih: Rp {saldo:,.0f}", 0, 1)

    filename = f"laporan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    pdf.output(filename)
    return filename

# =============================================
# CONVERSATION STATES
# =============================================

PEMASUKAN_INPUT, PENGELUARAN_INPUT, LAPORAN_INPUT = range(3)

# =============================================
# HANDLERS
# =============================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Halo! Saya Bot Pengelola Kas.\n\n"
        "Gunakan perintah berikut:\n"
        "/pemasukan - Catat pemasukan\n"
        "/pengeluaran - Catat pengeluaran\n"
        "/saldo - Lihat saldo\n"
        "/riwayat - Lihat riwayat transaksi\n"
        "/laporan - Download laporan"
    )


# PEMASUKAN

async def pemasukan_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Silakan ketik nominal dan keterangan pemasukan.\nContoh:\n50000 Gaji Bulanan",
        reply_markup=ReplyKeyboardRemove()
    )
    return PEMASUKAN_INPUT

async def pemasukan_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        parts = update.message.text.strip().split(" ", 1)
        nominal = float(parts[0])
        keterangan = parts[1] if len(parts) > 1 else "-"
        simpan_transaksi("pemasukan", nominal, keterangan)
        await update.message.reply_text(
            f"✅ Pemasukan Rp {nominal:,.0f} berhasil dicatat!\nKeterangan: {keterangan}"
        )
    except Exception:
        await update.message.reply_text("Format salah. Contoh:\n50000 Gaji Bulanan")

    return ConversationHandler.END

# PENGELUARAN

async def pengeluaran_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Silakan ketik nominal dan keterangan pengeluaran.\nContoh:\n20000 Beli Kopi",
        reply_markup=ReplyKeyboardRemove()
    )
    return PENGELUARAN_INPUT

async def pengeluaran_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        parts = update.message.text.strip().split(" ", 1)
        nominal = float(parts[0])
        keterangan = parts[1] if len(parts) > 1 else "-"
        simpan_transaksi("pengeluaran", nominal, keterangan)
        await update.message.reply_text(
            f"✅ Pengeluaran Rp {nominal:,.0f} berhasil dicatat!\nKeterangan: {keterangan}"
        )
    except Exception:
        await update.message.reply_text("Format salah. Contoh:\n20000 Beli Kopi")

    return ConversationHandler.END

# SALDO
async def saldo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    saldo = hitung_saldo()
    await update.message.reply_text(
        f"💰 *Saldo kas saat ini:* Rp {saldo:,}".replace(",", "."),
        parse_mode="Markdown"
    )

# RIWAYAT

async def riwayat_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    riwayat = get_riwayat(limit=10)

    if not riwayat:
        await update.message.reply_text("📭 Belum ada transaksi yang tercatat.")
        return

    pesan = "*📋 Riwayat Transaksi Terbaru:*\n\n"
    for i, (tanggal, jenis, nominal, keterangan) in enumerate(riwayat, start=1):
        jenis_ikon = "➕" if jenis == "pemasukan" else "➖"
        pesan += f"{i}. {jenis_ikon} *{jenis.title()}* - Rp {nominal:,}".replace(",", ".") + f"\n   📅 {tanggal}\n   📝 {keterangan}\n\n"

    await update.message.reply_text(pesan, parse_mode="Markdown")

# LAPORAN

async def laporan_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Ingin laporan semua transaksi atau bulan tertentu?\n"
        "Ketik:\nsemua\natau misalnya:\n07-2025"
    )
    return LAPORAN_INPUT

async def laporan_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    periode = update.message.text.strip().lower()

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    if periode == "semua":
        cursor.execute("SELECT tanggal, jenis, nominal, keterangan FROM transaksi ORDER BY tanggal")
        transaksi = cursor.fetchall()
        judul = "Semua Transaksi"
    else:
        try:
            bulan, tahun = periode.split("-")
            cursor.execute("""
                SELECT tanggal, jenis, nominal, keterangan
                FROM transaksi
                WHERE strftime('%m', tanggal) = ? AND strftime('%Y', tanggal) = ?
                ORDER BY tanggal
            """, (bulan.zfill(2), tahun))
            transaksi = cursor.fetchall()
            judul = f"Transaksi Bulan {periode}"
        except:
            await update.message.reply_text("Format salah. Ketik misalnya:\n07-2025")
            conn.close()
            return ConversationHandler.END

    conn.close()

    if not transaksi:
        await update.message.reply_text(f"Tidak ada transaksi untuk periode {periode}")
        return ConversationHandler.END

    await update.message.reply_text("⏳ Sedang membuat laporan...")

    filename = generate_pdf(transaksi, judul)

    with open(filename, "rb") as f:
        await update.message.reply_document(f, caption=f"Laporan Keuangan - {judul}")

    os.remove(filename)

    return ConversationHandler.END

# HANDLE TEKS DARI TOMBOL
async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text
    if text == "📥 Catat Pemasukan":
        return await pemasukan_start(update, context)
    elif text == "📤 Catat Pengeluaran":
        return await pengeluaran_start(update, context)
    elif text == "💰 Cek Saldo":
        saldo = hitung_saldo()
        await update.message.reply_text(f"💰 Saldo saat ini: Rp {saldo:,.0f}")
    elif text == "📋 Lihat Riwayat":
        riwayat = get_riwayat()
        if not riwayat:
            await update.message.reply_text("Belum ada transaksi.")
        else:
            message = "📋 Riwayat Transaksi:\n"
            for tgl, jenis, nominal, ket in riwayat:
                symbol = "+" if jenis == "pemasukan" else "-"
                message += f"{tgl} {symbol}{nominal:,.0f} {ket}\n"
            await update.message.reply_text(message)
    elif text == "📄 Download Laporan":
        return await laporan_start(update, context)
    else:
        await update.message.reply_text("Perintah tidak dikenal.")

# =============================================
# MAIN
# =============================================

def main():
    init_db()
    TOKEN = "7277923826:AAHuzYpFZ6y655eJN9ud76PpLP3SPELRhm8"

    app = ApplicationBuilder().token(TOKEN).build()

    conv_pemasukan = ConversationHandler(
        entry_points=[CommandHandler("pemasukan", pemasukan_start)],
        states={
            PEMASUKAN_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, pemasukan_input)],
        },
        fallbacks=[],
    )

    conv_pengeluaran = ConversationHandler(
        entry_points=[CommandHandler("pengeluaran", pengeluaran_start)],
        states={
            PENGELUARAN_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, pengeluaran_input)],
        },
        fallbacks=[],
    )

    conv_laporan = ConversationHandler(
        entry_points=[CommandHandler("laporan", laporan_start)],
        states={
            LAPORAN_INPUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, laporan_input)],
        },
        fallbacks=[],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_pemasukan)
    app.add_handler(conv_pengeluaran)
    app.add_handler(conv_laporan)
    app.add_handler(CommandHandler("saldo", saldo_handler))
    app.add_handler(CommandHandler("riwayat", riwayat_handler))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    print("Bot berjalan...")
    app.run_polling()

if __name__ == "__main__":
    main()
