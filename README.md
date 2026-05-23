# -*- coding: utf-8 -*-
"""
Dinamik Pivot Oluşturucu v13.5
- Otomatik Oturum Kaydı (Her adımda)
- Ayarları Temizle Butonu (Onaylı)
- Sabit Background.png (Transparan uyumlu)
- EXE Dönüştürme Uyumluluğu Eklendi
- CSV ve XLSX Okuma/Kaydetme Desteği (Hata Giderildi)
"""
import pandas as pd
import openpyxl
import xlsxwriter # EXE çıktısı için kütüphane bağımlılığı eklendi
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Listbox, END, MULTIPLE, Scrollbar
from tkinter import ttk
import json
import os
import sys

# --- GÖRSEL KÜTÜPHANESİ ---
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# --- Global Değişkenler ---
COL_WEIGHTS = [3, 3, 3, 1, 2, 1] 
header = ["Satırlar (Index)", "Sütunlar (Columns)", "Değerler (Values)", "İşlem (Seçiniz)", "Sayfa Adı", "Sil"]

pivot_entries = [] 
loaded_columns = []
bg_label_widget = None 

# --- DOSYA YOLU BULUCU (EXE UYUMLU) ---
def get_app_path():
    """Exe veya script'in çalıştığı klasörü bulur."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

DEFAULT_CONFIG_FILE = os.path.join(get_app_path(), "son_oturum.json")

# --- 1. SÜTUN YÜKLEME VE SEÇİCİ ---

def load_source_headers():
    global loaded_columns
    file_path = filedialog.askopenfilename(
        title="Sütun İsimlerini Okumak İçin Örnek Dosya Seçin", 
        filetypes=[("Veri Dosyaları", "*.xlsx *.xls *.csv"), ("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
    )
    if not file_path: return

    try:
        if file_path.lower().endswith('.csv'):
            # CSV ayırıcıyı otomatik bul (sep=None) ve Türkçe karakter desteği (utf-8-sig)
            df = pd.read_csv(file_path, sep=None, engine='python', nrows=0, encoding='utf-8-sig')
        else:
            df = pd.read_excel(file_path, sheet_name=0, nrows=0)
            
        loaded_columns = list(df.columns)
        
        lbl_status.config(text=f"✅ {len(loaded_columns)} adet sütun hafızaya alındı. Kutucuklara çift tıklayın!", fg="green")
        messagebox.showinfo("Başarılı", "Sütunlar yüklendi!\nKutucuklara ÇİFT TIKLAYARAK seçim yapabilirsiniz.")
    except Exception as e:
        messagebox.showerror("Hata", f"Dosya okunamadı: {e}")

def open_multi_selector(entry_widget, title):
    if not loaded_columns:
        messagebox.showwarning("Dikkat", "Henüz sütun listesi yüklenmedi.\nLütfen önce '1. ADIM: Sütunları Yükle' butonuna basınız.")
        return

    top = Toplevel()
    top.title(f"Seçim Yap: {title}")
    top.geometry("350x450")
    
    current_text = entry_widget.get()
    selected_vals = [x.strip() for x in current_text.split(',') if x.strip()]

    frame_list = tk.Frame(top)
    frame_list.pack(fill="both", expand=True, padx=10, pady=10)
    
    scrollbar = Scrollbar(frame_list)
    scrollbar.pack(side="right", fill="y")
    
    lb = Listbox(frame_list, selectmode=MULTIPLE, yscrollcommand=scrollbar.set, font=("Segoe UI", 10), height=15)
    lb.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=lb.yview)

    for idx, col in enumerate(loaded_columns):
        lb.insert(END, col)
        if col in selected_vals:
            lb.selection_set(idx)

    def save_selection():
        indices = lb.curselection()
        selected_items = [lb.get(i) for i in indices]
        entry_widget.delete(0, END)
        entry_widget.insert(0, ", ".join(selected_items))
        top.destroy()
        save_last_session() 

    tk.Button(top, text="Seçimi Kaydet ve Kapat", bg="#4CAF50", fg="white", font=("Segoe UI", 10, "bold"), command=save_selection).pack(fill="x", padx=10, pady=10)

# --- AYAR VE ŞABLON YÖNETİMİ ---

def get_config_data():
    data = []
    for row in pivot_entries:
        try:
            index, columns, values, aggfunc, sheet, _ = row
            data.append({
                "index": index.get(),
                "columns": columns.get(),
                "values": values.get(),
                "aggfunc": aggfunc.get(),
                "sheet": sheet.get()
            })
        except:
            continue
    full_config = {"rows": data}
    return full_config

def save_last_session():
    full_config = get_config_data()
    try:
        with open(DEFAULT_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(full_config, f, ensure_ascii=False, indent=4)
    except Exception:
        pass 

def clear_settings():
    if messagebox.askyesno("Ayarları Temizle", "Tüm ayarlar silinecek ve ekran sıfırlanacak.\nEmin misiniz?"):
        clear_all_rows()
        add_pivot_row() 
        save_last_session()

def save_template_as():
    config = get_config_data()
    if not config["rows"]: return
    row_data = config["rows"]

    file_path = filedialog.asksaveasfilename(
        title="Şablonu Farklı Kaydet",
        defaultextension=".json",
        filetypes=[("JSON Dosyası", "*.json"), ("Excel Dosyası", "*.xlsx")]
    )
    if not file_path: return

    try:
        if file_path.endswith(".xlsx"):
            df = pd.DataFrame(row_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Başarılı", "Şablon EXCEL formatında kaydedildi.")
        else:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("Başarılı", "Şablon JSON formatında kaydedildi.")
        save_last_session()
    except Exception as e:
        messagebox.showerror("Hata", f"Kaydedilemedi: {str(e)}")

def load_template_file():
    file_path = filedialog.askopenfilename(
        title="Şablon Yükle",
        filetypes=[("Tüm Dosyalar", "*.json *.xlsx"), ("JSON", "*.json"), ("Excel", "*.xlsx")]
    )
    if not file_path: return

    try:
        data_rows = []
        if file_path.endswith(".xlsx"):
            df = pd.read_excel(file_path).fillna("")
            data_rows = df.to_dict(orient='records')
        else:
            with open(file_path, "r", encoding="utf-8") as f:
                loaded_json = json.load(f)
                if isinstance(loaded_json, dict) and "rows" in loaded_json:
                    data_rows = loaded_json["rows"]
                elif isinstance(loaded_json, list):
                    data_rows = loaded_json
                else:
                    data_rows = []

        clear_all_rows()
        for item in data_rows:
            add_pivot_row(
                str(item.get("index", "")),
                str(item.get("columns", "")),
                str(item.get("values", "")),
                str(item.get("aggfunc", "count")),
                str(item.get("sheet", ""))
            )
        save_last_session()
        messagebox.showinfo("Başarılı", "Şablon yüklendi.")
    except Exception as e:
        messagebox.showerror("Hata", f"Yüklenemedi: {str(e)}")

# --- LOGO / BACKGROUND FONKSİYONLARI ---

def check_default_background():
    default_bg_path = os.path.join(get_app_path(), "background.png")
    if os.path.exists(default_bg_path) and PIL_AVAILABLE:
        update_background(default_bg_path)

def update_background(path):
    global bg_label_widget, photo_img
    try:
        img = Image.open(path)
        base_height = 70
        w_percent = (base_height / float(img.size[1]))
        w_size = int((float(img.size[0]) * float(w_percent)))
        if img.size[1] > base_height:
            img = img.resize((w_size, base_height), Image.Resampling.LANCZOS)
        photo_img = ImageTk.PhotoImage(img)
        if bg_label_widget:
            bg_label_widget.configure(image=photo_img)
        else:
            bg_label_widget = tk.Label(frm_bottom, image=photo_img, bg="#cfd8dc", bd=0)
        bg_label_widget.place(relx=1.0, rely=1.0, x=-10, y=-5, anchor="se")
    except Exception as e:
        print(f"Logo yükleme hatası: {e}")

# --- GUI SATIR YÖNETİMİ ---

def clear_all_rows():
    for row in pivot_entries:
        for w in row: w.destroy()
    pivot_entries.clear()

def add_pivot_row(index_val="", columns_val="", values_val="", aggfunc_val="count", sheet_val=""):
    i_entry = tk.Entry(scrollable_frame, font=("Arial", 10))
    c_entry = tk.Entry(scrollable_frame, font=("Arial", 10))
    v_entry = tk.Entry(scrollable_frame, font=("Arial", 10))
    a_entry = ttk.Combobox(scrollable_frame, values=["count", "sum", "mean", "min", "max", "nunique"], font=("Arial", 10), state="readonly")
    s_entry = tk.Entry(scrollable_frame, font=("Arial", 10))
    d_btn = tk.Button(scrollable_frame, text="Sil", bg="#ffcdd2", font=("Arial", 9))
    
    i_entry.insert(0, index_val)
    c_entry.insert(0, columns_val)
    v_entry.insert(0, values_val)
    if aggfunc_val in a_entry['values']:
        a_entry.set(aggfunc_val)
    else:
        a_entry.set("count")
    if not sheet_val: sheet_val = f"Analiz_{len(pivot_entries)+1}"
    s_entry.insert(0, sheet_val)

    i_entry.bind("<Double-Button-1>", lambda event, e=i_entry: open_multi_selector(e, "Satırlar"))
    i_entry.bind("<FocusOut>", lambda event: save_last_session())
    c_entry.bind("<Double-Button-1>", lambda event, e=c_entry: open_multi_selector(e, "Sütunlar"))
    c_entry.bind("<FocusOut>", lambda event: save_last_session())
    v_entry.bind("<Double-Button-1>", lambda event, e=v_entry: open_multi_selector(e, "Değerler"))
    v_entry.bind("<FocusOut>", lambda event: save_last_session())
    s_entry.bind("<FocusOut>", lambda event: save_last_session())
    a_entry.bind("<<ComboboxSelected>>", lambda event: save_last_session())

    widgets = [i_entry, c_entry, v_entry, a_entry, s_entry, d_btn]
    pivot_entries.append(widgets)
    d_btn.config(command=lambda w=widgets: delete_pivot_row(w))
    refresh_grid()
    save_last_session()

def delete_pivot_row(row_widgets):
    if len(pivot_entries) > 1 and messagebox.askyesno("Sil", "Silinsin mi?"):
        for w in row_widgets: w.destroy()
        pivot_entries.remove(row_widgets)
        refresh_grid()
        save_last_session()

def refresh_grid():
    for r, widgets in enumerate(pivot_entries):
        for c, w in enumerate(widgets):
            w.grid(row=r, column=c, padx=2, pady=2, sticky="ew")

# --- PROCESS: RAPORLAMA ---

def run_process():
    full_conf = get_config_data()
    configs = full_conf["rows"]
    valid_configs = [c for c in configs if c["index"] or c["values"]]
    if not valid_configs:
        messagebox.showwarning("Uyarı", "Lütfen ayar giriniz.")
        return
    save_last_session()
    
    file_paths = filedialog.askopenfilenames(
        title="Raporlanacak Dosyaları Seç", 
        filetypes=[("Veri Dosyaları", "*.xlsx *.xls *.csv"), ("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")]
    )
    
    if not file_paths: return
    count = 0
    errors = []
    for path in file_paths:
        try:
            is_csv = path.lower().endswith('.csv')
            if is_csv:
                # CSV Okuma (Hata Giderildi: sep=None ve encoding eklendi)
                df = pd.read_csv(path, sep=None, engine='python', encoding='utf-8-sig')
                original_sheet_name = "Veri_Seti"
                output_path = os.path.splitext(path)[0] + "_Pivot.xlsx"
            else:
                # Excel Okuma
                xl = pd.ExcelFile(path)
                df = pd.read_excel(path, sheet_name=0)
                original_sheet_name = xl.sheet_names[0]
                output_path = path

            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name=original_sheet_name, index=False)
                wb = writer.book
                fmt_header_left = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'left', 'valign': 'vcenter', 'indent': 1})
                fmt_header_center = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                fmt_cell = wb.add_format({'border': 1, 'align': 'left'})
                fmt_num = wb.add_format({'border': 1, 'num_format': '#,##0.00'})
                fmt_blue_num = wb.add_format({'bg_color': '#DDEBF7', 'border': 1, 'num_format': '#,##0.00', 'bold': True})
                fmt_blue_str = wb.add_format({'bg_color': '#DDEBF7', 'border': 1, 'bold': True, 'align': 'left'})
                fmt_pink_pct = wb.add_format({'bg_color': '#F2DCDB', 'border': 1, 'num_format': '0.00%', 'bold': True})
                
                for cfg in valid_configs:
                    idx = [x.strip() for x in cfg["index"].split(',') if x.strip()]
                    col = [x.strip() for x in cfg["columns"].split(',') if x.strip()]
                    val = [x.strip() for x in cfg["values"].split(',') if x.strip()]
                    if not val and not idx: continue
                    try:
                        piv = pd.pivot_table(df, index=idx, columns=col, values=val[0] if len(val)==1 else val, aggfunc=cfg["aggfunc"], fill_value=0)
                        if isinstance(piv.columns, pd.MultiIndex):
                            piv.columns = [' - '.join(map(str, c)).strip() for c in piv.columns.values]
                        piv["Genel Toplam"] = piv.sum(axis=1, numeric_only=True)
                        grand_total = piv["Genel Toplam"].sum()
                        piv["Yüzde Payı"] = piv["Genel Toplam"] / grand_total if grand_total != 0 else 0
                        bottom_row = piv.sum(numeric_only=True)
                        bottom_row.name = "Genel Toplam"
                        bottom_row["Yüzde Payı"] = 1.0 
                        piv = pd.concat([piv, bottom_row.to_frame().T])
                        sh_name = cfg["sheet"][:30]
                        ws = wb.add_worksheet(sh_name)
                        piv_flat = piv.reset_index()
                        rows, cols = piv_flat.shape
                        index_col_count = len(idx)
                        for c_idx, col_name in enumerate(piv_flat.columns):
                            if c_idx < index_col_count: ws.write(0, c_idx, str(col_name), fmt_header_left)
                            else: ws.write(0, c_idx, str(col_name), fmt_header_center)
                        for r in range(rows):
                            for c in range(cols):
                                value = piv_flat.iloc[r, c]
                                excel_row = r + 1
                                is_bottom = (r == rows - 1)
                                is_percent = (c == cols - 1)
                                is_total = (c == cols - 2)
                                is_num = isinstance(value, (int, float))
                                fmt = fmt_cell
                                if is_percent: fmt = fmt_pink_pct
                                elif is_total or is_bottom: fmt = fmt_blue_num if is_num else fmt_blue_str
                                elif is_num: fmt = fmt_num
                                ws.write(excel_row, c, value, fmt)
                        ws.set_column(0, index_col_count-1, 25)
                        ws.set_column(index_col_count, cols-3, 15)
                        ws.set_column(cols-2, cols-2, 18)
                        ws.set_column(cols-1, cols-1, 12)
                        data_col_count = cols - index_col_count - 2
                        if data_col_count > 0:
                            chart = wb.add_chart({'type': 'column'})
                            chart.set_style(10)
                            for i in range(data_col_count):
                                col_idx = index_col_count + i
                                chart.add_series({
                                    'name': [sh_name, 0, col_idx],
                                    # rows-1 kullanılarak en alttaki Genel Toplam satırı grafik dışı bırakıldı
                                    'categories': [sh_name, 1, 0, rows-1, 0],
                                    'values': [sh_name, 1, col_idx, rows-1, col_idx],
                                    'gap': 30
                                })
                            chart.set_title({'name': f'{sh_name} Grafiği'})
                            chart.set_size({'width': 600, 'height': 350})
                            ws.insert_chart(1, cols + 1, chart)
                    except Exception as e: print(f"Hata: {e}")
            count += 1
        except Exception as e: errors.append(f"{os.path.basename(path)}: {e}")
    msg = f"Tamamlandı: {count}"
    if errors: msg += f"\nHatalar:\n{errors}"
    messagebox.showinfo("Sonuç", msg)

# --- ANA EKRAN TASARIMI ---

root = tk.Tk()
root.title("Pivot Sihirbazı v13.5 (Oto-Kayıt & Temizleme)")
root.geometry("1100x750")

def on_closing():
    save_last_session()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

# MENÜ
frm_menu = tk.Frame(root, bg="#eee", bd=1, relief="raised")
frm_menu.pack(side="top", fill="x")
tk.Button(frm_menu, text="💾 Şablonu Kaydet", command=save_template_as, bg="white").pack(side="left", padx=5, pady=5)
tk.Button(frm_menu, text="📂 Şablon Yükle", command=load_template_file, bg="white").pack(side="left", padx=5, pady=5)

# GÖVDE
master_container = tk.Frame(root)
master_container.pack(fill="both", expand=True, padx=10, pady=10)

frm_header_container = tk.Frame(master_container)
frm_header_container.grid(row=0, column=0, sticky="ew")

frm_canvas = tk.Frame(master_container)
frm_canvas.grid(row=1, column=0, sticky="nsew")

scrollbar = tk.Scrollbar(master_container, orient="vertical")
scrollbar.grid(row=1, column=1, sticky="ns")

dummy_spacer = tk.Frame(master_container, width=17) 
dummy_spacer.grid(row=0, column=1, sticky="ns")

master_container.grid_columnconfigure(0, weight=1)
master_container.grid_rowconfigure(1, weight=1)

for i, h in enumerate(header):
    lbl = tk.Label(frm_header_container, text=h, font=("Arial", 10, "bold"), bg="#4472C4", fg="white", relief="raised", pady=5)
    lbl.grid(row=0, column=i, sticky="ew", padx=1)
    frm_header_container.grid_columnconfigure(i, weight=COL_WEIGHTS[i])

canvas = tk.Canvas(frm_canvas, bg="white", highlightthickness=0, yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.config(command=canvas.yview)

scrollable_frame = tk.Frame(canvas, bg="white")
scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

def on_canvas_configure(event):
    canvas.itemconfig(canvas_window, width=event.width)
canvas.bind("<Configure>", on_canvas_configure)

for i, w in enumerate(COL_WEIGHTS):
    scrollable_frame.grid_columnconfigure(i, weight=w)

# --- ALT BAR ---
frm_bottom = tk.Frame(root, bg="#cfd8dc", pady=10)
frm_bottom.pack(side="bottom", fill="x")

tk.Label(frm_bottom, text="Kutucuklara ÇİFT TIKLAYARAK sütun seçebilirsiniz.", bg="#cfd8dc", fg="#0277bd", font=("Arial", 9, "italic")).pack(side="top")
lbl_status = tk.Label(frm_bottom, text="⚠️ Önce 'Veri Yükle' yapınız.", bg="#cfd8dc", fg="red", font=("Arial", 9, "bold"))
lbl_status.pack(side="top", pady=2)

frm_mid_btns = tk.Frame(frm_bottom, bg="#cfd8dc")
frm_mid_btns.pack(side="top", pady=2)

tk.Button(frm_mid_btns, text="➕ Yeni Satır", command=lambda: add_pivot_row(), bg="white").pack(side="left", padx=5)
tk.Button(frm_mid_btns, text="🗑️ Ayarları Temizle", command=clear_settings, bg="#ef5350", fg="white", font=("Arial", 9, "bold")).pack(side="left", padx=5)

frm_btns = tk.Frame(frm_bottom, bg="#cfd8dc")
frm_btns.pack(pady=10) 

btn_load = tk.Button(frm_btns, text="1. Sütunları Yükle", command=load_source_headers, bg="#ffcc80", font=("Arial", 10, "bold"), padx=10, pady=5)
btn_load.pack(side="left", padx=10)
btn_run = tk.Button(frm_btns, text="2. Rapor & Grafik Oluştur", command=run_process, bg="#66bb6a", fg="white", font=("Arial", 11, "bold"), padx=10, pady=5)
btn_run.pack(side="left", padx=10)

def update_spacer(event):
    width = scrollbar.winfo_width()
    if width > 1: dummy_spacer.config(width=width)
scrollbar.bind("<Configure>", update_spacer)

# --- BAŞLANGIÇTA YÜKLEME ---
check_default_background()

if os.path.exists(DEFAULT_CONFIG_FILE):
    try:
        with open(DEFAULT_CONFIG_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
            rows_to_load = d.get("rows", []) if isinstance(d, dict) else (d if isinstance(d, list) else [])
            if not rows_to_load:
                add_pivot_row()
            else:
                for it in rows_to_load:
                    add_pivot_row(str(it.get("index","")), str(it.get("columns","")), str(it.get("values","")), str(it.get("aggfunc","count")), str(it.get("sheet","")))
    except:
        add_pivot_row()
else:
    add_pivot_row()

root.mainloop()