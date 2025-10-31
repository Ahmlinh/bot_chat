import os
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Konfigurasi
SCOPES = ['https://www.googleapis.com/auth/drive', 
          'https://www.googleapis.com/auth/spreadsheets']

# KONFIGURASI
SPREADSHEET_ID = '1lJa_d4SOTmLu_TqI8rSbjaKJaBw8i5cBYRYOT3TcYB8'  # Ganti dengan ID spreadsheet Anda
DRIVE_FOLDER_ID = None  # Akan dibuat otomatis

# Inisialisasi services
drive_service = None
sheets_service = None

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    
    # Load credentials
    SERVICE_ACCOUNT_FILE = 'credentials.json'
    
    if os.path.exists(SERVICE_ACCOUNT_FILE):
        credentials = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        
        # Inisialisasi services
        drive_service = build('drive', 'v3', credentials=credentials)
        sheets_service = build('sheets', 'v4', credentials=credentials)
        
        print("✅ Google API services berhasil diinisialisasi")
        
        # Buat folder otomatis di My Drive
        DRIVE_FOLDER_ID = create_drive_folder()
    else:
        print("❌ File credentials.json tidak ditemukan")
        
except Exception as e:
    print(f"❌ Error inisialisasi: {e}")

# Struktur kategori
KATEGORI = {
    'Surat Masuk': ['01 - Surat dari Perusahaan', '02 - Surat dari Pemerintah', '03 - Surat dari Individu'],
    'Surat Keluar': ['01 - Surat ke Perusahaan', '02 - Surat ke Pemerintah', '03 - Surat ke Individu'],
    'Keuangan': ['01 - Laporan Keuangan', '02 - Budget', '03 - Invoice'],
    'Kegiatan': ['01 - Rencana Kegiatan', '02 - Laporan Kegiatan', '03 - Foto Dokumentasi'],
    'Laporan': ['01 - Laporan Bulanan', '02 - Laporan Tahunan', '03 - Laporan Khusus']
}

def create_drive_folder(folder_name="Arsip Digital"):
    """Buat folder di Google Drive (My Drive)"""
    if not drive_service:
        print("❌ Drive service tidak tersedia")
        return None
    
    try:
        # Cek apakah folder sudah ada
        query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        
        existing_folders = drive_service.files().list(
            q=query,
            fields='files(id, name)'
        ).execute()
        
        if existing_folders.get('files'):
            folder_id = existing_folders['files'][0]['id']
            print(f"✅ Folder '{folder_name}' sudah ada: {folder_id}")
            return folder_id
        
        # Buat folder baru
        folder_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        
        folder = drive_service.files().create(
            body=folder_metadata,
            fields='id'
        ).execute()
        
        folder_id = folder.get('id')
        print(f"✅ Folder '{folder_name}' berhasil dibuat: {folder_id}")
        
        # Set permission agar file bisa diakses publik (view only)
        permission_metadata = {
            'type': 'anyone',
            'role': 'reader'
        }
        
        drive_service.permissions().create(
            fileId=folder_id,
            body=permission_metadata
        ).execute()
        
        print("✅ Permission public reader diberikan ke folder")
        return folder_id
        
    except Exception as e:
        print(f"❌ Error creating drive folder: {e}")
        return None

def upload_to_drive(file_path, file_name):
    """Upload file ke Google Drive folder"""
    if not drive_service or not DRIVE_FOLDER_ID:
        print("❌ Drive service atau folder ID tidak tersedia")
        return None, None
    
    try:
        file_metadata = {
            'name': file_name,
            'parents': [DRIVE_FOLDER_ID]
        }
        
        media = MediaFileUpload(file_path, resumable=True)
        
        print(f"🔄 Mengupload {file_name} ke Google Drive...")
        
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink, webContentLink'
        ).execute()
        
        file_id = file.get('id')
        web_view_link = file.get('webViewLink')
        
        print(f"✅ File berhasil diupload:")
        print(f"   File ID: {file_id}")
        print(f"   Link: {web_view_link}")
        
        return file_id, web_view_link
        
    except Exception as e:
        print(f"❌ Error uploading to Drive: {e}")
        return None, None

def save_file_locally(file, filename):
    """Simpan file secara lokal (fallback)"""
    try:
        upload_folder = 'static/uploads'
        if not os.path.exists(upload_folder):
            os.makedirs(upload_folder)
        
        file_extension = os.path.splitext(filename)[1]
        unique_filename = f"{secrets.token_hex(8)}{file_extension}"
        file_path = os.path.join(upload_folder, unique_filename)
        
        file.save(file_path)
        
        file_url = f"/static/uploads/{unique_filename}"
        print(f"✅ File disimpan secara lokal: {file_url}")
        return unique_filename, file_url
        
    except Exception as e:
        print(f"❌ Error saving file locally: {e}")
        return None, None

def get_next_archive_number(tingkat1, tingkat2):
    """Mendapatkan nomor arsip berikutnya"""
    if not sheets_service or SPREADSHEET_ID == 'your-spreadsheet-id-here':
        # Mode demo - generate dari data lokal
        all_archives = get_all_archives()
        max_num = 0
        
        for archive in all_archives:
            try:
                if archive['tingkat1'] == tingkat1:
                    nomor_arsip = archive['nomor_arsip']
                    num_part = nomor_arsip.split('.')[0]
                    num = int(num_part)
                    max_num = max(max_num, num)
            except (ValueError, IndexError):
                continue
        
        next_num = max_num + 1
        return f"{next_num:03d}"
    
    try:
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range='A2:H'
        ).execute()
        
        values = result.get('values', [])
        
        max_num = 0
        for row in values:
            if len(row) >= 3 and row[2] == tingkat1:
                try:
                    nomor_arsip = row[1]
                    num_part = nomor_arsip.split('.')[0]
                    num = int(num_part)
                    max_num = max(max_num, num)
                except (ValueError, IndexError):
                    continue
        
        next_num = max_num + 1
        return f"{next_num:03d}"
    except Exception as e:
        print(f"❌ Error getting next archive number: {e}")
        import random
        return f"{random.randint(1, 999):03d}"

def save_to_spreadsheet(data):
    """Simpan metadata ke Google Spreadsheet"""
    if not sheets_service or SPREADSHEET_ID == 'your-spreadsheet-id-here':
        # Mode demo - simpan ke file lokal
        return save_to_local_storage(data)
    
    try:
        values = [
            [
                data['id'],
                data['nomor_arsip'],
                data['tingkat1'],
                data['tingkat2'],
                data['judul'],
                data['deskripsi'],
                data['tanggal_upload'],
                data['link_drive'],
                data['nama_file']
            ]
        ]
        
        body = {
            'values': values
        }
        
        result = sheets_service.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range='A1',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()
        
        print(f"✅ Metadata berhasil disimpan ke spreadsheet")
        return True
    except Exception as e:
        print(f"❌ Error saving to spreadsheet: {e}")
        # Fallback ke local storage
        return save_to_local_storage(data)

def save_to_local_storage(data):
    """Simpan data ke file JSON lokal (fallback)"""
    try:
        storage_file = 'data/archives.json'
        
        # Buat folder data jika belum ada
        os.makedirs('data', exist_ok=True)
        
        # Load data existing
        archives = []
        if os.path.exists(storage_file):
            with open(storage_file, 'r', encoding='utf-8') as f:
                archives = json.load(f)
        
        # Tambah data baru
        archives.append(data)
        
        # Simpan ke file
        with open(storage_file, 'w', encoding='utf-8') as f:
            json.dump(archives, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Data disimpan secara lokal: {data['nomor_arsip']}")
        return True
    except Exception as e:
        print(f"❌ Error saving to local storage: {e}")
        return False

def get_all_archives():
    """Ambil semua data arsip"""
    # Coba ambil dari Google Sheets dulu
    if sheets_service and SPREADSHEET_ID != 'your-spreadsheet-id-here':
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=SPREADSHEET_ID,
                range='A2:I'
            ).execute()
            
            values = result.get('values', [])
            archives = []
            
            for row in values:
                if len(row) >= 9:
                    archive = {
                        'id': row[0],
                        'nomor_arsip': row[1],
                        'tingkat1': row[2],
                        'tingkat2': row[3],
                        'judul': row[4],
                        'deskripsi': row[5],
                        'tanggal_upload': row[6],
                        'link_drive': row[7],
                        'nama_file': row[8]
                    }
                    archives.append(archive)
            
            return archives
        except Exception as e:
            print(f"❌ Error getting archives from spreadsheet: {e}")
    
    # Fallback ke local storage
    try:
        storage_file = 'data/archives.json'
        if os.path.exists(storage_file):
            with open(storage_file, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"❌ Error loading local archives: {e}")
    
    # Data demo sebagai fallback terakhir
    return [
        {
            'id': 'demo1',
            'nomor_arsip': '001.01',
            'tingkat1': 'Surat Masuk',
            'tingkat2': '01 - Surat dari Perusahaan',
            'judul': 'Contoh Surat Masuk - DEMO',
            'deskripsi': 'Ini adalah data demo. File disimpan secara lokal.',
            'tanggal_upload': '2024-01-15 10:30:00',
            'link_drive': '#',
            'nama_file': 'demo_surat_masuk.pdf'
        }
    ]

# Routes (sama seperti sebelumnya)
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        try:
            tingkat1 = request.form['tingkat1']
            tingkat2 = request.form['tingkat2']
            judul = request.form['judul']
            deskripsi = request.form['deskripsi']
            file = request.files['file']
            
            if file and file.filename != '':
                # Validasi file size
                file.seek(0, os.SEEK_END)
                file_size = file.tell()
                file.seek(0)
                
                if file_size > 10 * 1024 * 1024:
                    return jsonify({
                        'success': False, 
                        'message': 'File terlalu besar. Maksimal 10MB.'
                    })
                
                # Generate nomor arsip
                nomor_arsip_tingkat1 = get_next_archive_number(tingkat1, tingkat2)
                tingkat2_number = tingkat2.split(' - ')[0]
                nomor_arsip = f"{nomor_arsip_tingkat1}.{tingkat2_number}"
                
                # Coba upload ke Google Drive dulu
                file_id, drive_link = None, None
                if drive_service and DRIVE_FOLDER_ID:
                    # Simpan file sementara untuk upload
                    temp_path = f"temp_{secrets.token_hex(8)}_{file.filename}"
                    file.save(temp_path)
                    
                    try:
                        file_id, drive_link = upload_to_drive(temp_path, file.filename)
                    finally:
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
                
                # Jika gagal upload ke Drive, simpan lokal
                if not file_id or not drive_link:
                    print("🔄 Menggunakan penyimpanan lokal")
                    saved_filename, drive_link = save_file_locally(file, file.filename)
                    file_id = saved_filename
                
                if file_id and drive_link:
                    archive_data = {
                        'id': secrets.token_hex(8),
                        'nomor_arsip': nomor_arsip,
                        'tingkat1': tingkat1,
                        'tingkat2': tingkat2,
                        'judul': judul,
                        'deskripsi': deskripsi,
                        'tanggal_upload': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'link_drive': drive_link,
                        'nama_file': file.filename
                    }
                    
                    if save_to_spreadsheet(archive_data):
                        return jsonify({
                            'success': True, 
                            'message': f'Arsip berhasil diupload! Nomor Arsip: {nomor_arsip}',
                            'nomor_arsip': nomor_arsip
                        })
                    else:
                        return jsonify({
                            'success': False, 
                            'message': 'File berhasil diupload tetapi gagal menyimpan metadata.'
                        })
                else:
                    return jsonify({
                        'success': False, 
                        'message': 'Gagal mengupload file. Coba lagi.'
                    })
            
            return jsonify({'success': False, 'message': 'File tidak valid!'})
            
        except Exception as e:
            print(f"❌ Error in upload: {e}")
            return jsonify({'success': False, 'message': f'Terjadi kesalahan: {str(e)}'})
    
    return render_template('upload.html', kategori=KATEGORI)

@app.route('/arsip')
def arsip():
    archives = get_all_archives()
    return render_template('arsip.html', archives=archives)

@app.route('/search', methods=['GET', 'POST'])
def search():
    if request.method == 'POST':
        keyword = request.form.get('keyword', '').lower()
        tingkat1 = request.form.get('tingkat1', '')
        tingkat2 = request.form.get('tingkat2', '')
        
        all_archives = get_all_archives()
        filtered_archives = []
        
        for archive in all_archives:
            match = True
            
            if keyword:
                keyword_match = (
                    keyword in archive['judul'].lower() or 
                    keyword in archive['nomor_arsip'].lower() or
                    keyword in archive['deskripsi'].lower()
                )
                if not keyword_match:
                    match = False
            
            if tingkat1 and archive['tingkat1'] != tingkat1:
                match = False
                
            if tingkat2 and archive['tingkat2'] != tingkat2:
                match = False
                
            if match:
                filtered_archives.append(archive)
        
        return render_template('search.html', archives=filtered_archives, 
                             kategori=KATEGORI, search_performed=True)
    
    return render_template('search.html', kategori=KATEGORI, search_performed=False)

@app.route('/api/kategori/<tingkat1>')
def get_subkategori(tingkat1):
    subkategori = KATEGORI.get(tingkat1, [])
    return jsonify(subkategori)

@app.route('/health')
def health():
    """Endpoint untuk mengecek status sistem"""
    status = {
        'drive_service': 'available' if drive_service else 'unavailable',
        'sheets_service': 'available' if sheets_service else 'unavailable',
        'spreadsheet_id': 'configured' if SPREADSHEET_ID != 'your-spreadsheet-id-here' else 'not configured',
        'drive_folder_id': 'available' if DRIVE_FOLDER_ID else 'not available',
        'credentials_file': 'found' if os.path.exists('credentials.json') else 'not found',
        'mode': 'production' if (drive_service and sheets_service and DRIVE_FOLDER_ID) else 'demo'
    }
    return jsonify(status)

if __name__ == '__main__':
    print("🚀 Aplikasi Arsip Digital")
    print("=" * 50)
    
    # Cek konfigurasi
    if not os.path.exists('credentials.json'):
        print("❌ File credentials.json tidak ditemukan!")
    else:
        print("✅ File credentials.json ditemukan")
    
    if SPREADSHEET_ID == 'your-spreadsheet-id-here':
        print("⚠️  SPREADSHEET_ID belum dikonfigurasi")
    else:
        print("✅ SPREADSHEET_ID sudah dikonfigurasi")
    
    print(f"\n📁 Folder Drive: {'✅ Tersedia' if DRIVE_FOLDER_ID else '❌ Gagal dibuat'}")
    
    print("\n📋 Mode operasi:")
    if drive_service and sheets_service and DRIVE_FOLDER_ID:
        print("   🟢 PRODUCTION - Terhubung ke Google API")
    else:
        print("   🟡 DEMO - Menggunakan penyimpanan lokal")
    
    print(f"\n🌐 Aplikasi berjalan di http://localhost:5000")
    print("💡 Kunjungi /health untuk status detail")
    
    app.run(debug=True, host='0.0.0.0', port=5000)