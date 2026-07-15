import pandas as pd
from fpdf import FPDF
from tkinter import Tk, filedialog, messagebox, Button, Checkbutton, IntVar, Label
from tkinter.ttk import Progressbar
from PIL import Image, ImageTk
from concurrent.futures import ThreadPoolExecutor
import threading
import os
import sys
from datetime import datetime
from dateutil import parser


class MyFPDF(FPDF):
    def header(self):
        pass


class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("400x400")
        self.root.resizable(False, False)

        self.files = []  # Seçilen dosyalar
        self.stop_at_blank_var = IntVar()

        # Pencere ikonu ve logo
        self.root.iconbitmap(self.resource_path('assets/favicon.ico'))
        logo_path = self.resource_path('assets/image.png')
        logo_image = Image.open(logo_path)
        
        # Logo boyutunu pencereye uygun şekilde ölçeklendir (max genişlik 280, max yükseklik 80)
        max_w, max_h = 280, 80
        w, h = logo_image.size
        ratio = min(max_w / w, max_h / h)
        logo_image = logo_image.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
        
        logo_photo = ImageTk.PhotoImage(logo_image)
        logo_label = Label(self.root, image=logo_photo)
        logo_label.image = logo_photo
        logo_label.pack(pady=10)

        # Kaç dosya seçildiğini gösteren label
        self.file_count_label = Label(self.root, text="Seçilen dosya sayısı: 0", font=("Arial", 12))
        self.file_count_label.pack(pady=5)

        # Kontrol elemanları
        check_button = Checkbutton(self.root, text="Boş satırdan sonra dur", variable=self.stop_at_blank_var)
        check_button.pack(pady=10)

        # İlerleme çubuğu
        self.progress = Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Dosya seçme butonu
        select_button = Button(self.root, text="Dosya Seç", command=self.select_files)
        select_button.pack(pady=10)

        # Dönüştür butonu
        convert_button = Button(self.root, text="Dönüştür", command=self.start_conversion)
        convert_button.pack(pady=10)

    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)


    def format_date_value(self, val):
        if pd.isna(val) or val == "":
            return ""
        # If it's a datetime/Timestamp
        if hasattr(val, 'day') and hasattr(val, 'month') and hasattr(val, 'year'):
            return f"{val.day}.{val.month}.{val.year}"
        
        val_str = str(val).strip()
        # If it contains time portion, remove it (e.g., '2025-03-27 00:00:00' -> '2025-03-27')
        if " " in val_str:
            val_str = val_str.split(" ")[0]
            
        try:
            parsed = parser.parse(val_str, dayfirst=True, fuzzy=True)
            return f"{parsed.day}.{parsed.month}.{parsed.year}"
        except Exception:
            return val_str

    def format_time_value(self, val):
        if pd.isna(val) or val == "":
            return ""
        # If it is a datetime.time or Timestamp/datetime
        if hasattr(val, 'hour') and hasattr(val, 'minute') and hasattr(val, 'second'):
            if hasattr(val, 'microsecond') and val.microsecond > 0:
                ms_str = f"{val.microsecond // 100000:01d}" # one decimal digit
                return f"{val.hour:02d}:{val.minute:02d}:{val.second:02d}.{ms_str}"
            else:
                return f"{val.hour:02d}:{val.minute:02d}:{val.second:02d}"
                
        val_str = str(val).strip()
        # If it has date portion, remove it
        if " " in val_str:
            val_str = val_str.split(" ")[-1]
            
        parts = val_str.split(":")
        if len(parts) == 3:
            sec_parts = parts[2].split(".")
            if len(sec_parts) == 2:
                ms = sec_parts[1][:1] # keep 1 decimal place
                return f"{parts[0]:0>2}:{parts[1]:0>2}:{sec_parts[0]:0>2}.{ms}"
            else:
                return f"{parts[0]:0>2}:{parts[1]:0>2}:{parts[2]:0>2}"
        return val_str

    def normalize_date(self, value):
        """ Tarih formatını esnek şekilde normalize eder (d.M.yyyy) """
        return self.format_date_value(value)

    def process_file(self, excel_path, output_folder, stop_at_blank):
        """ Tek bir dosyayı işleyip PDF'e dönüştürür. Hata durumunda Exception fırlatır. """
        font_path = self.resource_path('assets/DejaVuSans.ttf')
        
        if str(excel_path).lower().endswith('.csv'):
            import csv
            import io
            encodings = ['utf-8', 'cp1254', 'latin1']
            df = None
            for enc in encodings:
                try:
                    # errors='replace' ve NUL byte temizliği ile okuma
                    with open(excel_path, 'r', encoding=enc, errors='replace') as f:
                        content = f.read().replace('\0', '')
                        
                    if not content.strip():
                        df = pd.DataFrame()
                        break
                        
                    sep = ';' if content.count(';') > content.count(',') else ','
                    reader = csv.reader(io.StringIO(content), delimiter=sep)
                    data = list(reader)
                    
                    if not data:
                        df = pd.DataFrame()
                        break
                    
                    max_cols = max(len(row) for row in data)
                    headers = data[0]
                    # Eğer başlıklar benzersiz değilse veya boşsa, isimlendirme çakışmalarını önleyelim
                    headers = [str(h) if h else f"Unnamed: {i}" for i, h in enumerate(headers)]
                    
                    if len(headers) < max_cols:
                        headers.extend([f"Unnamed: {i}" for i in range(len(headers), max_cols)])
                    
                    # Eksik hücreleri None ile doldur (Pandas için NaN olur)
                    padded_data = [row + [None]*(max_cols - len(row)) for row in data[1:]]
                    df = pd.DataFrame(padded_data, columns=headers)
                    break
                except Exception:
                    continue
            
            if df is None:
                # Tüm encoding denemeleri başarısız olursa python engine ile standart yolla devam et
                try:
                    df = pd.read_csv(excel_path, dtype=object, engine='python', on_bad_lines='skip', sep=None)
                except Exception:
                    df = pd.read_csv(excel_path, dtype=object, engine='python', on_bad_lines='skip', sep=None, encoding='cp1254')
        else:
            df = pd.read_excel(excel_path, dtype=object)  # Orijinal tipleri korumak için object olarak oku

        # Sütun tespiti (Tarih ve Saat kolonlarını bul)
        date_column = None
        time_column = None
        
        for col in df.columns:
            col_str = str(col).lower()
            if 'date' in col_str or 'tarih' in col_str:
                date_column = col
            elif 'time' in col_str or 'saat' in col_str:
                time_column = col
                
        if date_column is None and len(df.columns) > 0:
            date_column = df.columns[0]
        if time_column is None and len(df.columns) > 1:
            time_column = df.columns[1]

        # Tarih ve saat formatlarını temizle/normalize et
        if date_column in df.columns:
            df[date_column] = df[date_column].apply(self.format_date_value)
        if time_column in df.columns:
            df[time_column] = df[time_column].apply(self.format_time_value)

        # Ondalık sayılarda 1 hane koruma ve genel numerik formatlama (tarih/saat hariç)
        for col in df.columns:
            if col == date_column or col == time_column:
                continue
            def format_numeric(val):
                if pd.isna(val) or val == "":
                    return ""
                val_str_orig = str(val)
                val_normalized = val_str_orig.replace(',', '.')
                try:
                    val_float = float(val_normalized)
                    # Excel'deki görünüm ile eşleşmesi için 1 ondalık basamağa yuvarla
                    val_rounded = round(val_float, 1)
                    if val_rounded.is_integer():
                        return str(int(val_rounded))
                    
                    formatted = f"{val_rounded:.1f}"
                    if ',' in val_str_orig:
                        return formatted.replace('.', ',')
                    return formatted.replace('.', ',') # Türkiye/Avrupa standardı için her zaman virgül kullan
                except (ValueError, TypeError):
                    return val_str_orig
            df[col] = df[col].apply(format_numeric)

        # Target temperature column tespiti (ORTAM1 veya 1. sensör)
        temp_col = None
        for col in ['ORTAM1', '1', 1]:
            if col in df.columns:
                temp_col = col
                break
        if temp_col is None and len(df.columns) > 2:
            temp_col = df.columns[2]

        if stop_at_blank:
            if temp_col is None or temp_col not in df.columns:
                raise KeyError(f"Sütun bulunamadı: 'ORTAM1' veya sensör sütunu ('Boş satırdan sonra dur' seçeneği bu sütunu gerektirir)")
            for idx, row in df.iterrows():
                val = row[temp_col]
                if pd.isnull(val) or str(val).strip() == "":
                    df = df[:idx]
                    break

        # Probe (sensör) kolonlarını belirle (Tarih ve Saat kolonları ve isimsiz kolonlar hariç)
        probe_cols = [col for col in df.columns if col != date_column and col != time_column and not str(col).startswith("Unnamed:")]

        # Hedef hücreleri (her bir prop'un 56'ya ulaştığı ilk hücre) bul
        highlight_cells = set() # (row_idx, col_name)
        reached_probes = set()

        for i in range(len(df)):
            for col in probe_cols:
                if col not in reached_probes:
                    val = df[col].iloc[i]
                    try:
                        val_float = float(str(val).replace(',', '.'))
                        if val_float >= 56.0:
                            highlight_cells.add((i, col))
                            reached_probes.add(col)
                    except (ValueError, TypeError):
                        pass

        # Tüm propların 56'ya ulaşıp ulaşmadığını kontrol et
        all_reached = (len(reached_probes) == len(probe_cols)) and len(probe_cols) > 0

        # "START OF EXPOSURE" fallback satırını bul
        exposure_row_idx = None
        for i in range(len(df)):
            if any("START OF EXPOSURE" in str(df[item].iloc[i]) for item in df.columns):
                exposure_row_idx = i
                break

        # Dinamik yönlendirme (Sütun sayısı 8'den fazlaysa Yatay A4)
        orientation = 'L' if df.shape[1] > 8 else 'P'
        pdf = MyFPDF(orientation=orientation)
        pdf.add_page()
        pdf.add_font('DejaVu', '', font_path)
        pdf.set_font('DejaVu', '', 8)

        # Dinamik sütun genişliği hesaplama (overlapping engellemek için)
        # Margins: sol 10mm, sağ 10mm. Yazdırılabilir alan: pdf.w - 20
        printable_width = pdf.w - 20
        raw_widths = []
        for col in df.columns:
            # Sütundaki tüm hücrelerin ve başlığın genişliğini bul
            max_w = max([pdf.get_string_width(str(val)) for val in df[col]] + [pdf.get_string_width(str(col))])
            raw_widths.append(max_w + 4)  # 4mm boşluk bırak
            
        total_raw_width = sum(raw_widths)
        col_widths = [w * (printable_width / total_raw_width) for w in raw_widths]
        row_height = pdf.font_size

        # Sütun başlıklarını yazdır
        for col_idx, header in enumerate(df.columns):
            if str(header).startswith("Unnamed:"):  # İsimsiz kolonları boş bırak
                header = ""
            pdf.cell(col_widths[col_idx], row_height * 2, str(header), border=0)
        pdf.ln(row_height * 2)

        # Veri satırlarını yazdır
        for i in range(len(df)):
            # Sadece tüm proplar 56'ya ulaşmadıysa fallback exposure satırını vurgula
            is_exposure_row = False
            if not all_reached and exposure_row_idx is not None:
                if i == exposure_row_idx:
                    is_exposure_row = True

            for col_idx, item in enumerate(df.columns):
                value = df[item].iloc[i]
                if pd.isnull(value):
                    value = ""

                cell_highlight = False
                if is_exposure_row:
                    cell_highlight = True
                elif (i, item) in highlight_cells:
                    cell_highlight = True

                # Eğer hücre highlight edilecekse, fill=True ile yazdır (sarı arka plan)
                if cell_highlight:
                    pdf.set_fill_color(255, 255, 0)
                    pdf.cell(col_widths[col_idx], row_height * 2, str(value), border=0, fill=True)
                else:
                    pdf.cell(col_widths[col_idx], row_height * 2, str(value), border=0)

            pdf.ln(row_height * 2)  # Bir sonraki satıra geç

        base_name = os.path.splitext(os.path.basename(excel_path))[0]
        output_file = f"{output_folder}/{base_name}.pdf"
        pdf.output(output_file)

    def process_file_wrapper(self, excel_path, output_folder, stop_at_blank):
        """ process_file metodunu çağırır ve hata durumlarını yakalar """
        try:
            self.process_file(excel_path, output_folder, stop_at_blank)
            return True, ""
        except Exception as e:
            import traceback
            err_msg = str(e)
            print(f"Hata oluştu ({excel_path}): {err_msg}")
            traceback.print_exc()
            return False, err_msg

    def batch_process(self, files, output_folder):
        """ Çoklu dosyaları işlemek için ThreadPoolExecutor kullanır """
        total_files = len(files)
        self.progress['value'] = 0
        self.root.update()

        completed_files = 0
        failed_files = []

        futures = []
        with ThreadPoolExecutor() as executor:
            for file_path in files:
                future = executor.submit(self.process_file_wrapper, file_path, output_folder, self.stop_at_blank_var.get())
                futures.append((file_path, future))
                
            for idx, (file_path, future) in enumerate(futures):
                success, err_msg = future.result()
                if success:
                    completed_files += 1
                else:
                    failed_files.append((file_path, err_msg))
                
                self.progress['value'] = ((idx + 1) / total_files) * 100
                self.root.update()

        self.progress['value'] = 0
        self.root.update()

        if failed_files:
            error_details = "\n".join([f"- {os.path.basename(f)}: {err}" for f, err in failed_files])
            messagebox.showerror(
                "Dönüştürme Hatası", 
                f"Bazı dosyalar dönüştürülemedi:\n\n{error_details}"
            )
        else:
            messagebox.showinfo(
                "Başarılı", 
                f"{total_files} dosya başarıyla dönüştürüldü."
            )

    def select_files(self):
        """ Dosyaları seç ve dosya sayısını güncelle """
        self.files = filedialog.askopenfilenames(
            title="Excel veya CSV Dosyalarını Seç", 
            filetypes=[("Excel ve CSV dosyaları", "*.xlsx *.xls *.csv"), ("Tüm dosyalar", "*.*")]
        )
        self.file_count_label.config(text=f"Seçilen dosya sayısı: {len(self.files)}")

    def start_conversion(self):
        """ Dönüştürme işlemini başlatır """
        if not self.files:
            messagebox.showwarning("Uyarı", "Dönüştürmek için dosya seçilmedi!")
            return

        output_folder = filedialog.askdirectory(title="PDF'lerin Kaydedileceği Klasörü Seç")
        if output_folder:
            threading.Thread(target=self.batch_process, args=(self.files, output_folder)).start()
        else:
            messagebox.showwarning("Uyarı", "PDF'lerin kaydedileceği klasör seçilmedi.")


if __name__ == "__main__":
    root = Tk()
    app = PDFConverterApp(root)
    root.mainloop()

