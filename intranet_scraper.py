import requests
from bs4 import BeautifulSoup
import csv
import sys
import time
import re
import os
from requests.compat import urljoin
import html
import pandas as pd

# --- KONFIGURASI UMUM ---
BASE_URL = 'http://172.16.5.10/mypatologi/'

# --- KONFIGURASI LOGIN ---
LOGIN_URL = BASE_URL + 'index.php?redirect'
INTRANET_USERNAME = 'puguh'
INTRANET_PASSWORD = '987654'
LOGIN_FIELDS = {
    'username_field_name': 'field[]',
    'password_field_name': 'field[]',
    'button_x_field': 'x',
    'button_x_value': '54',
    'button_y_field': 'y',
    'button_y_value': '12'
}

# --- KONFIGURASI PENCARIAN ---
SEARCH_AJAX_URL = BASE_URL + 'ajax.php?m=patologi&s=register&a=cek'

# --- STRUKTUR DATA DAN HEADER ---
DISPLAY_HEADERS_STRUCTURED = {
    "DATA REGISTRASI": {
        "No. Medical Rec": 'No. Medical Rec',
        "Nama Pasien": 'Nama Pasien',
        "Tanggal Terima": 'Tanggal Terima',
        "Tanggal Bayar": 'Tanggal Bayar',
        "Tanggal Jawab": 'Tanggal Jawab',
        "Dokter Pengirim": 'Dokter Pengirim',
        "Jenis Kelamin": 'Jenis Kelamin',
        "Umur": 'Umur',
        "No Imun": 'No Imun',
        "No PA": 'No PA',
        "RS Asal / Bagian": 'RS Asal / Bagian',
    },
    "DATA AWAL": {
        "Riwayat / Diagnosis klinik": 'Riwayat / Diagnosis klinik',
        "Diagnosis PA": 'Diagnosis PA',
        "Kartu lama": 'Kartu lama',
    },
    "DATA HASIL IMUNOHISTOKIMIA": {
        "Hasil immunohistokimia": 'Hasil immunohistokimia',
        "Kesimpulan (IHK)": 'Diagnosis IHK',
        "Lanjutan": 'Lanjutan',
        "Keterangan": 'Keterangan',
    },
    "AKHIR": {
        "Kesimpulan (Akhir)": 'Kesimpulan Tambahan',
        "Anjuran": 'Anjuran Tambahan',
    },
    "Lain-lain (Tidak di Screenshot)": {
        "Topologi": 'Topologi',
        "Morfologi": 'Morfologi',
    }
}

ALL_POSSIBLE_HEADERS = [key_name for category in DISPLAY_HEADERS_STRUCTURED.values() for key_name in category.values()]

# --- FUNGSI UTILITY ---
def _get_text_or_empty(element, separator='\n'):
    """Mengekstrak teks dari elemen BeautifulSoup, mengembalikan 'kosong' jika tidak ada."""
    if element:
        text = element.get_text(separator=separator, strip=True).replace('\xa0', ' ').strip()
        return text if text else "kosong"
    return "kosong"

def _extract_view_button_onclick(button_onclick_attr):
    """Mengekstrak URL detail dari atribut onclick tombol 'VIEW'."""
    if not button_onclick_attr: return None
    match_replace = re.search(r"window\.location\.replace\(['\"](.*?)['\"]\)", button_onclick_attr)
    if match_replace:
        return urljoin(BASE_URL, html.unescape(match_replace.group(1)))
    match_preview = re.search(r"show_preview\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:,\s*(\d+))?\)", button_onclick_attr)
    if match_preview:
        lab, patient_id, id_lab, id_sublab = match_preview.groups('0')
        return f"{BASE_URL}event.php?m=patologi&s=dok&a=antri&lab={lab}&id={patient_id}&sub={id_lab}&subsub={id_sublab}"
    return None

# --- FUNGSI LOGIN ---
def login_to_intranet(session):
    """Melakukan login ke sistem intranet."""
    login_data = {
        LOGIN_FIELDS['username_field_name']: [INTRANET_USERNAME, INTRANET_PASSWORD],
        LOGIN_FIELDS['button_x_field']: LOGIN_FIELDS['button_x_value'],
        LOGIN_FIELDS['button_y_field']: LOGIN_FIELDS['button_y_value']
    }
    try:
        print(f"Mencoba login ke {LOGIN_URL}...")
        response = session.post(LOGIN_URL, data=login_data, timeout=30)
        response.raise_for_status()
        if "LOG OFF" in response.text or "HOME :: APLIKASI PATOLOGI" in response.text:
            print("Login berhasil.")
            return True
        print("Login gagal. Periksa kredensial atau struktur form login.")
        return False
    except requests.exceptions.RequestException as e:
        print(f"Error saat login: {e}")
        return False

# --- FUNGSI PENCARIAN & EKSTRAKSI HASIL TABEL ---
def extract_table_row_data(row_soup):
    """Mengekstrak data dari satu baris tabel hasil pencarian."""
    cells = row_soup.find_all('td')
    if len(cells) < 12: return None
    view_button = cells[11].find('input', type='button', value=lambda v: v and v.strip().lower() == 'view')
    return {'No.Reg': cells[0].text.strip(), 'Nama Pasien': cells[1].text.strip(), 'No Medical Record': cells[7].text.strip(), 'View Button Onclick': view_button.get('onclick') if view_button else ''}

def search_patient_by_pa_ihk(session, pa_ihk_input):
    """Melakukan pencarian pasien berdasarkan Nomor PA/IHK."""
    print(f"\n--- Melakukan pencarian untuk No PA/IHK = '{pa_ihk_input}' ---")
    search_data = {'keyword': pa_ihk_input, 'opt': '3', 'Submit': '   Cari   '}
    try:
        response = session.post(SEARCH_AJAX_URL, data=search_data, timeout=60)
        response.raise_for_status()
        return response.text
    except requests.exceptions.RequestException as e:
        print(f"Error saat pencarian untuk '{pa_ihk_input}': {e}")
        return None

# --- FUNGSI EKSTRAKSI DATA DARI HALAMAN DETAIL ---
def extract_data_from_detail_page(session, detail_url):
    """Mengunjungi halaman detail dan mengekstrak semua data yang dibutuhkan."""
    data = {header: "kosong" for header in ALL_POSSIBLE_HEADERS}
    print(f"Mengambil data dari: {detail_url}")

    try:
        response = session.get(detail_url, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        with open("debug_detail_page.html", "w", encoding="utf-8") as f:
            f.write(soup.prettify())
        print(f"--- DEBUG: HTML halaman detail disimpan ke 'debug_detail_page.html' ---")

        main_table = soup.find('table', class_='form')
        if not main_table:
            print("  -> ERROR: Tabel utama dengan class='form' tidak ditemukan. Ekstraksi dibatalkan.")
            return data

        print("\n--- Memulai Ekstraksi Data ---")
        
        # --- Ekstraksi DATA REGISTRASI (Metode Eksplisit) ---
        label_med_rec = main_table.find('td', string=re.compile(r'No\.\s*Medical\s*Rec'))
        if label_med_rec: data['No. Medical Rec'] = _get_text_or_empty(label_med_rec.find_next_sibling('td'))

        label_nama = main_table.find('td', string=re.compile(r'Nama\s*Pasien'))
        if label_nama: data['Nama Pasien'] = _get_text_or_empty(label_nama.find_next_sibling('td'))

        label_tgl_terima = main_table.find('td', string=re.compile(r'Tanggal\s*Terima'))
        if label_tgl_terima: data['Tanggal Terima'] = _get_text_or_empty(label_tgl_terima.find_next_sibling('td'))

        label_tgl_bayar = main_table.find('td', string=re.compile(r'Tanggal\s*Bayar'))
        if label_tgl_bayar: data['Tanggal Bayar'] = _get_text_or_empty(label_tgl_bayar.find_next_sibling('td'))
        
        label_tgl_jawab = main_table.find('td', string=re.compile(r'Tanggal\s*Jawab'))
        if label_tgl_jawab: data['Tanggal Jawab'] = _get_text_or_empty(label_tgl_jawab.find_next_sibling('td'))

        label_dokter = main_table.find('td', string=re.compile(r'Dokter\s*Pengirim'))
        if label_dokter: data['Dokter Pengirim'] = _get_text_or_empty(label_dokter.find_next_sibling('td'))
        
        label_rs = main_table.find('td', string=re.compile(r'RS\s*Asal\s*/\s*Bagian'))
        if label_rs: data['RS Asal / Bagian'] = _get_text_or_empty(label_rs.find_next_sibling('td'))

        label_jk_umur = main_table.find('td', string=re.compile(r'Jenis\s*Kelamin\s*/\s*Umur'))
        if label_jk_umur:
            jk_umur_raw = _get_text_or_empty(label_jk_umur.find_next_sibling('td'))
            if '/' in jk_umur_raw:
                parts = [p.strip() for p in jk_umur_raw.split('/', 1)]
                data['Jenis Kelamin'] = parts[0] or "kosong"
                data['Umur'] = parts[1] if len(parts) > 1 and parts[1] else "kosong"
            else: data['Jenis Kelamin'] = jk_umur_raw or "kosong"

        label_imun_pa = main_table.find('td', string=re.compile(r'No\s*Imun\s*/\s*No\s*PA'))
        if label_imun_pa:
            imun_pa_raw = _get_text_or_empty(label_imun_pa.find_next_sibling('td'))
            if '/' in imun_pa_raw:
                parts = [p.strip() for p in imun_pa_raw.split('/', 1)]
                data['No Imun'] = parts[0] or "kosong"
                data['No PA'] = parts[1] if len(parts) > 1 and parts[1] else "kosong"
            else: data['No PA'] = imun_pa_raw or "kosong"

        # --- Ekstraksi DATA AWAL ---
        label_riwayat = main_table.find('td', string=re.compile(r'Riwayat\s*/\s*Diagnosis\s*klinik'))
        if label_riwayat: data['Riwayat / Diagnosis klinik'] = _get_text_or_empty(label_riwayat.find_next_sibling('td'))

        label_diag_pa = main_table.find('td', string=re.compile(r'Diagnosis\s*PA'))
        if label_diag_pa: data['Diagnosis PA'] = _get_text_or_empty(label_diag_pa.find_next_sibling('td'))

        label_kartu = main_table.find('td', string=re.compile(r'Kartu\s*lama'))
        if label_kartu: data['Kartu lama'] = _get_text_or_empty(label_kartu.find_next_sibling('td'))

        # --- Ekstraksi DATA HASIL IMUNOHISTOKIMIA ---
        label_hasil_ihk = main_table.find('td', string=re.compile(r'Hasil\s*immunohistokimia'))
        if label_hasil_ihk: data['Hasil immunohistokimia'] = _get_text_or_empty(label_hasil_ihk.find_next_sibling('td'))

        label_lanjutan = main_table.find('td', string=re.compile(r'Lanjutan'))
        if label_lanjutan: data['Lanjutan'] = _get_text_or_empty(label_lanjutan.find_next_sibling('td'))

        label_kesimpulan_ihk = main_table.find('td', {'width': '175'}, string=re.compile(r'Kesimpulan'))
        if label_kesimpulan_ihk: data['Diagnosis IHK'] = _get_text_or_empty(label_kesimpulan_ihk.find_next_sibling('td'))

        label_keterangan_ihk = main_table.find('td', {'width': '175'}, string=re.compile(r'Keterangan'))
        if label_keterangan_ihk: data['Keterangan'] = _get_text_or_empty(label_keterangan_ihk.find_next_sibling('td'))

        # --- Ekstraksi Bagian Akhir (Kesimpulan, Anjuran, Morfologi, Topologi) ---
        all_tds = main_table.find_all('td')
        kesimpulan_labels, anjuran_labels, morfologi_labels, topologi_labels = [], [], [], []
        
        patterns = {
            'kesimpulan': (re.compile(r'^\s*Kesimpulan\s*:?\s*$', re.IGNORECASE), kesimpulan_labels),
            'anjuran': (re.compile(r'^\s*Anjuran\s*:?\s*$', re.IGNORECASE), anjuran_labels),
            'morfologi': (re.compile(r'^\s*Morfologi\s*:?\s*$', re.IGNORECASE), morfologi_labels),
            'topologi': (re.compile(r'^\s*Topologi\s*:?\s*$', re.IGNORECASE), topologi_labels),
        }

        for td in all_tds:
            if 'width' in td.attrs: continue # Abaikan sel dengan atribut width (seperti Kesimpulan IHK)
            td_text = td.get_text(strip=True)
            for key, (pattern, label_list) in patterns.items():
                if pattern.match(td_text) and td.parent.name == 'tr':
                    label_list.append(td)
                    break 

        if kesimpulan_labels:
            data['Kesimpulan Tambahan'] = _get_text_or_empty(kesimpulan_labels[-1].find_next_sibling('td'))
        if anjuran_labels:
            data['Anjuran Tambahan'] = _get_text_or_empty(anjuran_labels[-1].find_next_sibling('td'))
        if morfologi_labels:
            data['Morfologi'] = _get_text_or_empty(morfologi_labels[-1].find_next_sibling('td'))
        if topologi_labels:
            data['Topologi'] = _get_text_or_empty(topologi_labels[-1].find_next_sibling('td'))

        print("--- Ekstraksi Selesai ---")
        print("--- Ringkasan Data Ditemukan ---")
        for key, value in data.items():
            if value != "kosong":
                print(f"  {key}: {str(value)[:70]}...")
        print("------------------------------")

    except requests.exceptions.RequestException as e:
        print(f"Error saat mengambil halaman detail: {e}")
    except Exception as e:
        print(f"Terjadi kesalahan tidak terduga saat memproses halaman detail: {e}")

    return data

# --- FUNGSI UTAMA ---
def main():
    print("Pastikan Anda telah menginstal 'pandas' dan 'openpyxl': pip install pandas openpyxl")
    
    session = requests.Session()
    if not login_to_intranet(session):
        sys.exit("Script berhenti karena login gagal.")

    print("\n--- Input Data dari Excel ---")
    excel_file_path = input("Masukkan path lengkap file Excel (.xlsx): ").strip()
    if not os.path.exists(excel_file_path):
        print(f"Error: File '{excel_file_path}' tidak ditemukan.")
        sys.exit("Script berhenti.")

    sheet_name_or_index = input("Masukkan nama sheet atau nomor indeks (0 untuk sheet pertama): ").strip()
    try:
        sheet_to_read = int(sheet_name_or_index) if sheet_name_or_index.isdigit() else sheet_name_or_index
        df = pd.read_excel(excel_file_path, sheet_name=sheet_to_read, dtype=str).fillna('')
        print(f"Berhasil membaca {len(df)} baris dari sheet '{sheet_to_read}'.")
    except Exception as e:
        print(f"Error membaca file Excel: {e}")
        sys.exit("Script berhenti.")

    print("\n--- Pilih Rentang Baris Excel ---")
    try:
        start_row = int(input(f"Mulai dari baris (1 - {len(df)}): ") or 1)
        end_row = int(input(f"Sampai baris (1 - {len(df)}): ") or len(df))
        if not (1 <= start_row <= len(df) and start_row <= end_row <= len(df)): raise ValueError("Rentang tidak valid")
    except ValueError as e:
        print(f"Input tidak valid: {e}. Gunakan rentang default (semua baris).")
        start_row, end_row = 1, len(df)

    df_to_process = df.iloc[start_row-1:end_row]
    print(f"Akan memproses {len(df_to_process)} baris (baris Excel {start_row} sampai {end_row}).")

    print("\n--- Pemetaan Kolom Excel (Hardcoded) ---")
    pa_col_name, nama_col_name, medrec_col_name = 'No IHK', 'Nama Pasien', 'Med Rec'
    for col in [pa_col_name, nama_col_name, medrec_col_name]:
        if col not in df.columns:
            print(f"Error: Kolom '{col}' tidak ditemukan di file Excel Anda.")
            sys.exit("Script berhenti.")

    all_results = []
    print(f"\n--- Memulai Pemrosesan {len(df_to_process)} Baris dari Excel ---")

    for index, row in df_to_process.iterrows():
        print(f"\n==================================================")
        print(f"Memproses baris Excel #{index + 1}")
        
        pa_ihk_from_excel = str(row.get(pa_col_name, '')).strip()
        nama_from_excel = str(row.get(nama_col_name, '')).strip()
        medrec_from_excel = str(row.get(medrec_col_name, '')).strip()

        # Inisialisasi baris output dengan data dari Excel dan default value
        output_row = {header: "Tidak ada data" for header in ALL_POSSIBLE_HEADERS}
        output_row['No IHK (Excel)'] = pa_ihk_from_excel
        output_row['Nama Pasien (Excel)'] = nama_from_excel
        output_row['Med Rec (Excel)'] = medrec_from_excel
        
        if not pa_ihk_from_excel:
            print(f"  Peringatan: No IHK kosong. Melewati.")
            all_results.append(output_row)
            continue

        ajax_response_html = search_patient_by_pa_ihk(session, pa_ihk_from_excel)
        if not ajax_response_html or not BeautifulSoup(ajax_response_html, 'html.parser').find('table', class_='list'):
            print(f"  Peringatan: Tidak ditemukan hasil pencarian untuk No IHK '{pa_ihk_from_excel}'.")
            all_results.append(output_row)
            continue
            
        soup_results = BeautifulSoup(ajax_response_html, 'html.parser')
        data_rows = soup_results.find('table', class_='list').find_all('tr')[1:]
        print(f"  Ditemukan {len(data_rows)} baris di tabel hasil pencarian.")
        
        found_match = False
        for r_soup in data_rows:
            table_row_data = extract_table_row_data(r_soup)
            if not table_row_data: continue

            name_match = nama_from_excel.lower() in table_row_data['Nama Pasien'].lower().strip() if nama_from_excel else True
            medrec_match = medrec_from_excel == table_row_data['No Medical Record'].strip() if medrec_from_excel else True

            if name_match and medrec_match:
                print(f"  -> COCOK: Data Excel cocok dengan data web.")
                detail_url = _extract_view_button_onclick(table_row_data['View Button Onclick'])
                if detail_url:
                    extracted_data = extract_data_from_detail_page(session, detail_url)
                    output_row.update(extracted_data) # Update baris output dengan data yang ditemukan
                    print("  -> SUKSES: Data lengkap berhasil diekstrak.")
                else:
                    print("  -> GAGAL: Tidak dapat mengekstrak URL detail dari tombol VIEW.")
                found_match = True
                break # Hentikan pencarian setelah menemukan kecocokan (berhasil atau gagal view)
            
        if not found_match:
            print(f"  Peringatan: Tidak ada data yang cocok di intranet untuk baris Excel ini.")
        
        all_results.append(output_row)
        time.sleep(1)

    if not all_results:
        print("\nTidak ada data yang berhasil diproses. Proses selesai.")
        sys.exit()

    print(f"\n--- Total Baris Diproses: {len(all_results)} ---")
    
    # --- Pilihan Kolom Output ---
    print("\n--- Pilihan Kolom Output ---")
    selection_map = {}
    counter = 1
    for category, items in DISPLAY_HEADERS_STRUCTURED.items():
        print(f"\n--- {category} ---")
        for display_name, key_name in items.items():
            print(f"{counter}. {display_name}")
            selection_map[str(counter)] = key_name
            counter += 1
    
    selected_indices_str = input("\nMasukkan nomor data yang diinginkan (pisahkan koma) atau kosongkan untuk semua: ").strip()
    
    selected_keys = [selection_map[n.strip()] for n in selected_indices_str.split(',') if n.strip() in selection_map]
    
    if not selected_keys:
        print("Tidak ada kolom valid yang dipilih. Menampilkan semua kolom.")
        selected_keys = ALL_POSSIBLE_HEADERS

    # --- Output ke Excel ---
    output_excel_file_name = input("\nMasukkan nama file Excel untuk menyimpan hasil (cth: hasil_data.xlsx): ").strip() or "output_data_intranet.xlsx"
    if not output_excel_file_name.lower().endswith('.xlsx'): output_excel_file_name += '.xlsx'

    try:
        # Buat DataFrame dari semua hasil
        df_output = pd.DataFrame(all_results)
        
        # Siapkan kolom untuk output, mulai dari kolom sumber lalu kolom pilihan
        final_columns = ['No IHK (Excel)', 'Nama Pasien (Excel)', 'Med Rec (Excel)'] + selected_keys
        
        # Pastikan semua kolom ada di DataFrame, jika tidak ada tambahkan kolom kosong
        for col in final_columns:
            if col not in df_output.columns:
                df_output[col] = ''
        
        df_output = df_output[final_columns]
        
        df_output.to_excel(output_excel_file_name, index=False, engine='openpyxl')
        print(f"\nData telah disimpan ke '{os.path.abspath(output_excel_file_name)}'.")

    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan ke Excel: {e}")

    print("\nProses selesai.")

if __name__ == '__main__':
    main()
