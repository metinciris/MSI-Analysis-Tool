import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import warnings
import sys

def extract_msi_data(file_path):
    warnings.simplefilter(action='ignore', category=UserWarning)  # OpenPyXL uyarısını gizle
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Sample name alımı
    sample_name_raw = df.iloc[2, 1]  # Örnek: "MP7-25_S73_L001_R1_001_Mapped Reads"
    sample_name = sample_name_raw.split("_")[0]  # "MP7-25"
    
    # Clinical term alımı
    clinical_term = df.iloc[9, 1]  # Örnek: "MS-stable"
    
    # Tablo verilerini al
    table_start_index = 11
    df_table = df.iloc[table_start_index + 1:, :7]
    df_table.columns = ["Locus", "Coverage", "Read count", "Baseline lengths", 
                        "Coverage ratio", "Stability threshold", "Stability"]
    df_table = df_table.reset_index(drop=True)
    
    # Stability sütunundaki Unstable olanları bul
    unstable_loci = df_table[df_table["Stability"] == "Unstable"]["Locus"].tolist()
    
    # MSI Durumu Belirleme (Yeni format)
    if len(unstable_loci) == 0:
        msi_status = "- Mikrosatellit Stabil (MS-Stable/MSS). Dokuz bölgenin hiç birinde instabilite saptanmadı."
    elif len(unstable_loci) == 1:
        msi_status = f"- Mikrosatellit Stabil (MS-Stable/MSS). Dokuz bölgenin birinde ({unstable_loci[0]}) instabilite saptandı."
    else:
        count_map = {2: "ikisinde", 3: "üçünde", 4: "dördünde", 5: "beşinde", 6: "altısında", 7: "yedisinde", 8: "sekizinde", 9: "dokuzunda"}
        count_word = count_map.get(len(unstable_loci), f"{len(unstable_loci)} bölgesinde")
        
        if len(unstable_loci) >= 4:
            msi_status = "- Mikrosatellit High (MS-High/MSS). "
        elif len(unstable_loci) >= 2:
            msi_status = "- Mikrosatellit Low (MS-Low/MSS). "
        else:
            msi_status = "- Mikrosatellit Stabil (MS-Stable/MSS). "
        
        msi_status += f"Dokuz bölgenin {count_word} ({', '.join(unstable_loci)}) instabilite saptandı."
    
    return f"{sample_name}  {clinical_term}", msi_status

def copy_to_clipboard(text, window):
    window.clipboard_clear()
    window.clipboard_append(text)
    window.update()

def close_and_exit(window):
    window.destroy()
    sys.exit()

def select_files():
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("Bilgi", "MSI Rapor Analizi Başlatılıyor. Lütfen MSI sonuçlarını içeren Excel dosyalarınızı toplu halde seçin.")
    
    file_paths = filedialog.askopenfilenames(title="MSI Rapor Excel Dosyalarını Seçin", 
                                             filetypes=[("Excel files", "*.xlsx")])
    return file_paths

def main():
    file_paths = select_files()
    
    results = []
    for file in file_paths:
        sample, status = extract_msi_data(file)
        results.append(f"{sample}\n{status}\n")
    
    result_text = "\n".join(results)
    
    display_window = tk.Tk()
    display_window.title("MSI Analiz Sonuçları")
    display_window.geometry("1200x800")
    
    text_box = tk.Text(display_window, wrap=tk.WORD, font=("Arial", 18), padx=20, pady=20)
    text_box.insert(tk.END, result_text)
    text_box.pack(expand=True, fill=tk.BOTH)
    text_box.config(state=tk.NORMAL)
    
    button_frame = tk.Frame(display_window)
    button_frame.pack(pady=10)
    
    copy_button = tk.Button(button_frame, text="Tümünü Kopyala", font=("Arial", 16), command=lambda: copy_to_clipboard(result_text, display_window))
    copy_button.pack(side=tk.LEFT, padx=20)
    
    close_button = tk.Button(button_frame, text="Kapat", font=("Arial", 16), command=lambda: close_and_exit(display_window))
    close_button.pack(side=tk.RIGHT, padx=20)
    
    display_window.mainloop()
    
if __name__ == "__main__":
    main()
