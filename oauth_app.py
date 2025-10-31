import os
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, session, redirect, url_for
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Konfigurasi OAuth 2.0
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]

# Konfigurasi - GANTI DENGAN DATA ANDA
CLIENT_SECRETS_FILE = "client_secrets.json"
SPREADSHEET_ID = "your-spreadsheet-id"
DRIVE_FOLDER_ID = "your-folder-id"  # Bisa menggunakan folder personal atau shared drive

# Struktur kategori
KATEGORI = {
    'Surat Masuk': ['01 - Surat dari Perusahaan', '02 - Surat dari Pemerintah', '03 - Surat dari Individu'],
    'Surat Keluar': ['01 - Surat ke Perusahaan', '02 - Surat ke Pemerintah', '03 - Surat ke Individu'],
    'Keuangan': ['01 - Laporan Keuangan', '02 - Budget', '03 - Invoice'],
    'Kegiatan': ['01 - Rencana Kegiatan', '02 - Laporan Kegiatan', '03 - Foto Dokumentasi'],
    'Laporan': ['01 - Laporan Bulanan', '02 - Laporan Tahunan', '03 - Laporan Khusus']
}

def get_google_services():
    """Mendapatkan Google services dengan OAuth credentials"""
    creds = None
    
    # Load token dari session
    if 'credentials' in session:
        creds = Credentials.from_authorized_user_info(json.loads(session['credentials']), SCOPES)
    
    # Jika credentials tidak valid, request login ulang
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            session['credentials'] = json.dumps({
                'token': creds.token,
                'refresh_token': creds.refresh_token,
                'token_uri': creds.token_uri,
                'client_id': creds.client_id,
                'client_secret': creds.client_secret,
                'scopes': creds.scopes
            })
        else:
            return None, None
    
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        return drive_service, sheets_service
    except Exception as e:
        print(f"Error building services: {e}")
        return None, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/login')
def login():
    """Endpoint untuk login dengan Google OAuth"""
    # Create flow instance to manage the OAuth 2.0 Authorization Grant Flow steps.
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for('oauth2callback', _external=True)
    )

    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )

    # Store the state so the callback can verify the auth server response.
    session['state'] = state

    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    """Callback untuk OAuth 2.0"""
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        state=session['state'],
        redirect_uri=url_for('oauth2callback', _external=True)
    )

    flow.fetch_token(authorization_response=request.url)

    # Store credentials in session
    credentials = flow.credentials
    session['credentials'] = json.dumps({
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    })

    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    """Logout dan hapus credentials"""
    session.pop('credentials', None)
    session.pop('state', None)
    return redirect(url_for('index'))

def is_authenticated():
    """Cek apakah user sudah login"""
    return 'credentials' in session

def get_next_archive_number(tingkat1, tingkat2):
    """Mendapatkan nomor arsip berikutnya berdasarkan tingkat"""
    drive_service, sheets_service = get_google_services()
    
    if not sheets_service:
        import random
        return f"{random.randint(1, 999):03d}"
    
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
        print(f"Error getting next archive number: {e}")
        import random
        return f"{random.randint(1, 999):03d}"

def upload_to_drive(file_path, file_name):
    """Upload file ke Google Drive"""
    drive_service, _ = get_google_services()
    
    if not drive_service:
        return None, None
    
    try:
        file_metadata = {
            'name': file_name
        }
        
        if DRIVE_FOLDER_ID and DRIVE_FOLDER_ID != 'your-folder-id':
            file_metadata['parents'] = [DRIVE_FOLDER_ID]
        
        media = MediaFileUpload(file_path, resumable=True)
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        return file.get('id'), file.get('webViewLink')
    except Exception as e:
        print(f"Error uploading to Drive: {e}")
        return None, None

def save_to_spreadsheet(data):
    """Simpan metadata ke Google Spreadsheet"""
    _, sheets_service = get_google_services()
    
    if not sheets_service:
        print(f"📝 Simulasi save: {data['nomor_arsip']} - {data['judul']}")
        return True
    
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
        
        return True
    except Exception as e:
        print(f"Error saving to spreadsheet: {e}")
        return False

def get_all_archives():
    """Ambil semua data arsip dari spreadsheet"""
    _, sheets_service = get_google_services()
    
    if not sheets_service:
        return [
            {
                'id': 'demo1',
                'nomor_arsip': '001.01',
                'tingkat1': 'Surat Masuk',
                'tingkat2': '01 - Surat dari Perusahaan',
                'judul': 'Surat Penawaran Kerjasama',
                'deskripsi': 'Surat penawaran kerjasama dari PT Contoh',
                'tanggal_upload': '2024-01-15 10:30:00',
                'link_drive': 'https://drive.google.com/file/d/demo1/view',
                'nama_file': 'surat_penawaran.pdf'
            }
        ]
    
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
        print(f"Error getting archives: {e}")
        return []

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if not is_authenticated():
        return jsonify({'success': False, 'message': 'Silakan login terlebih dahulu'})
    
    if request.method == 'POST':
        try:
            tingkat1 = request.form['tingkat1']
            tingkat2 = request.form['tingkat2']
            judul = request.form['judul']
            deskripsi = request.form['deskripsi']
            file = request.files['file']
            
            if file and file.filename != '':
                file_path = f"temp_{secrets.token_hex(8)}_{file.filename}"
                file.save(file_path)
                
                try:
                    nomor_arsip_tingkat1 = get_next_archive_number(tingkat1, tingkat2)
                    tingkat2_number = tingkat2.split(' - ')[0]
                    nomor_arsip = f"{nomor_arsip_tingkat1}.{tingkat2_number}"
                    
                    file_id, drive_link = upload_to_drive(file_path, file.filename)
                    
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
                            if os.path.exists(file_path):
                                os.remove(file_path)
                            
                            return jsonify({
                                'success': True, 
                                'message': f'Arsip berhasil diupload! Nomor Arsip: {nomor_arsip}',
                                'nomor_arsip': nomor_arsip
                            })
                
                finally:
                    if os.path.exists(file_path):
                        os.remove(file_path)
            
            return jsonify({'success': False, 'message': 'Gagal mengupload arsip!'})
            
        except Exception as e:
            print(f"Error in upload: {e}")
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
def health_check():
    status = {
        'authenticated': is_authenticated(),
        'spreadsheet_id': 'configured' if SPREADSHEET_ID != 'your-spreadsheet-id' else 'not configured',
        'drive_folder_id': 'configured' if DRIVE_FOLDER_ID != 'your-folder-id' else 'not configured'
    }
    return jsonify(status)

if __name__ == '__main__':
    print("🚀 Aplikasi Arsip Digital dengan OAuth 2.0")
    print("🌐 Berjalan di http://localhost:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)