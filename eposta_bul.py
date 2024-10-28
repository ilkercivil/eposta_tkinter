import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# E-posta adresi bulmak için regex (grup eklenmiş haliyle)
email_regex = r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'

def excel_dosyalarini_sec():
    """Birden fazla Excel dosyası seçmek için dosya seçim penceresi."""
    dosya_yollari = filedialog.askopenfilenames(
        title="Excel Dosyalarını Seçin",
        filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
    )
    if dosya_yollari:
        dosya_yolu_label.config(text=f"{len(dosya_yollari)} dosya seçildi")
        eposta_bul(dosya_yollari)

def eposta_bul(dosya_yollari):
    """Seçilen Excel dosyalarından e-posta adreslerini bulur ve ekrana yazar."""
    email_listesi = []

    try:
        for dosya_yolu in dosya_yollari:
            df = pd.read_excel(dosya_yolu, sheet_name=None)

            # Her sayfadaki tabloyu tarar
            for sayfa_adi, tablo in df.items():
                for sutun in tablo.columns:
                    tablo[sutun] = tablo[sutun].astype(str)  # Tip dönüşümü
                    bulunan_mailler = tablo[sutun].str.extractall(email_regex)
                    email_listesi.extend(bulunan_mailler[0].tolist())

        # Bulunan e-postaları ekrana yazdır
        email_listesi = list(set(email_listesi))  # Tekilleştir
        eposta_goster(email_listesi)

    except Exception as e:
        messagebox.showerror("Hata", f"Dosyalar okunurken bir hata oluştu: {e}")

def eposta_goster(email_listesi):
    """Bulunan e-posta adreslerini ekrandaki text alanına yazar."""
    eposta_text.config(state=tk.NORMAL)
    eposta_text.delete(1.0, tk.END)  # Önceki metni temizle
    if email_listesi:
        for email in email_listesi:
            eposta_text.insert(tk.END, email + "\n")
    else:
        eposta_text.insert(tk.END, "Hiç e-posta adresi bulunamadı.\n")
    eposta_text.config(state=tk.DISABLED)

# Tkinter arayüzünü oluşturma
pencere = tk.Tk()
pencere.title("Excel Dosyalarından E-posta Adresi Bulma")
pencere.geometry("600x400")

# Dosya seçme butonu
dosya_sec_butonu = tk.Button(pencere, text="Excel Dosyalarını Seç", command=excel_dosyalarini_sec)
dosya_sec_butonu.pack(pady=10)

# Seçilen dosya sayısını göstermek için label
dosya_yolu_label = tk.Label(pencere, text="Henüz bir dosya seçilmedi.", wraplength=500)
dosya_yolu_label.pack(pady=5)

# E-posta adreslerini göstermek için scrolled text alanı
eposta_text = scrolledtext.ScrolledText(pencere, width=70, height=15, state=tk.DISABLED)
eposta_text.pack(pady=10)

# Pencereyi çalıştır
pencere.mainloop()
