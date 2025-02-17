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


    def normalize_date(self, value):
        """ Tarih formatını esnek şekilde normalize eder (d.M.yyyy) """
        if pd.isna(value) or value == "":
            return ""
        try:
            # Sayısal bir değer ise (Excel tarih serial formatı)
            if isinstance(value, (int, float)):
                date = datetime.fromordinal(datetime(1900, 1, 1).toordinal()) + int(value) - 2
                return date.strftime("%-d.%-m.%Y")
            # String ise parse et
            elif isinstance(value, str):
                # Tarih ve saat kısmını ayır
                value = value.split()[0]
                parsed_date = parser.parse(value, dayfirst=True, fuzzy=True)
                return parsed_date.strftime("%-d.%-m.%Y")
        except Exception as e:
            print(f"Geçersiz tarih formatı: {value} - {str(e)}")
        return str(value)  # Fallback olarak orijinal değeri string olarak döndür

    def process_file(self, excel_path, output_folder, stop_at_blank):
        """ Tek bir dosyayı işleyip PDF'e dönüştürür """
        font_path = self.resource_path('assets/DejaVuSans.ttf')
        try:
            df = pd.read_excel(excel_path, dtype=str)  # Tüm verileri string olarak oku

            # İlk sütun adını al
            time_column = df.columns[0]

            # Tarih formatlarını normalize et
            for col in df.columns:
                df[col] = df[col].apply(self.normalize_date)

            # Saat formatlarını temizle
            def clean_time(value):
                if isinstance(value, str):
                    if len(value.split(":")) == 2:
                        return value
                    elif len(value.split(":")) == 3:
                        return ":".join(value.split(":")[:2])
                return value

            df[time_column] = df[time_column].apply(clean_time)

            # Ondalık sayılarda iki hane koruma
            numeric_columns = df.select_dtypes(include=['float64', 'int64']).columns
            for col in numeric_columns:
                df[col] = df[col].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "")

            if stop_at_blank:
                for idx, row in df.iterrows():
                    if pd.isnull(row['ORTAM1']):
                        df = df[:idx]
                        break

            pdf = MyFPDF()
            pdf.add_page()
            pdf.add_font('DejaVu', '', font_path)
            pdf.set_font('DejaVu', '', 8)

            col_width = pdf.w / (df.shape[1] + 1)
            row_height = pdf.font_size

            # Sütun başlıklarını yazdır
            for header in df.columns:
                if header.startswith("Unnamed:"):  # İsimsiz kolonları boş bırak
                    header = ""
                pdf.cell(col_width, row_height * 2, str(header), border=0)
            pdf.ln(row_height * 2)

            # Veri satırlarını yazdır
            for i in range(len(df)):
                is_highlighted = False  # Sarı vurgulama için flag

                # Eğer "START OF EXPOSURE" varsa bütün satırı highlight et
                if any("START OF EXPOSURE" in str(df[item].iloc[i]) for item in df.columns):
                    is_highlighted = True

                for item in df.columns:
                    value = df[item].iloc[i]
                    if pd.isnull(value):
                        value = ""

                    # Eğer satır highlight edilecekse, fill=True ile yazdır
                    if is_highlighted:
                        pdf.set_fill_color(255, 255, 0)  # Sarı arka plan
                        pdf.cell(col_width, row_height * 2, str(value), border=0, fill=True)
                    else:
                        pdf.cell(col_width, row_height * 2, str(value), border=0)

                pdf.ln(row_height * 2)  # Bir sonraki satıra geç

            output_file = f"{output_folder}/{os.path.basename(excel_path).replace('.xlsx', '.pdf')}"
            pdf.output(output_file)

        except Exception as e:
            print(f"Hata oluştu: {excel_path} - {e}")

    def batch_process(self, files, output_folder):
        """ Çoklu dosyaları işlemek için ThreadPoolExecutor kullanır """
        total_files = len(files)
        self.progress['value'] = 0
        self.root.update()

        completed_files = 0

        with ThreadPoolExecutor() as executor:
            for file_path in files:
                executor.submit(self.process_file, file_path, output_folder, self.stop_at_blank_var.get())
                completed_files += 1
                self.progress['value'] = (completed_files / total_files) * 100
                self.root.update()

        self.progress['value'] = 0
        self.root.update()
        messagebox.showinfo("Başarılı", f"{total_files} dosya seçildi, {completed_files} dosya başarıyla dönüştürüldü.")

    def select_files(self):
        """ Dosyaları seç ve dosya sayısını güncelle """
        self.files = filedialog.askopenfilenames(title="Excel Dosyalarını Seç", filetypes=[("Excel files", "*.xlsx *.xls")])
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

