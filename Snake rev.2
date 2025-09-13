#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
import os
import sys
import pandas as pd
import numpy as np
import glob
import re

# ---------- Utils ----------
def resource_path(relative_path):
    """取得開發/打包後通用的資源路徑"""
    try:
        base_path = sys._MEIPASS  # PyInstaller onefile 的暫存解壓路徑
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def extract_timestamp_ms(filepath):
    """從檔名抽出第一段連續數字當作毫秒時間戳"""
    name = os.path.splitext(os.path.basename(filepath))[0]
    m = re.search(r'(\d+)', name)
    return int(m.group(1)) if m else None

def generate_time_points(start_s, time1, time2, time3, interval1, interval2, interval3):
    """
    產生秒數清單：
      0 ~ time1     每 interval1 秒
      20 ~ time2    每 interval2 秒
      90 ~ time3    每 interval3 秒
    然後整體再加上 start_s 偏移
    """
    # 轉 int，避免 range() 出錯
    time1 = int(time1); time2 = int(time2); time3 = int(time3)
    interval1 = max(1, int(interval1))
    interval2 = max(1, int(interval2))
    interval3 = max(1, int(interval3))
    start_s = float(start_s)

    offsets = []
    offsets += list(range(0,   time1 + 1, interval1))  # 0~time1
    offsets += list(range(20,  time2 + 1, interval2))  # 20~time2
    offsets += list(range(90,  time3 + 1, interval3))  # 90~time3

    # 去重 & 排序（避免交疊造成重複時間點）
    offsets = sorted(set(offsets))
    return [start_s + off for off in offsets]

def read_spectrum(filepath, wl_min, wl_max):
    """
    讀 dat：假設為 TSV，第二欄 WL、第三欄 R
    只保留波長 [wl_min, wl_max]
    """
    df = pd.read_csv(filepath, sep='\t', skiprows=1, names=['PIXEL','WL','R'])
    mask = (df['WL'] >= wl_min) & (df['WL'] <= wl_max)
    wl = df.loc[mask, 'WL'].values
    refl = df.loc[mask, 'R'].values
    return wl, refl

def process_folder(data_dir, start_s, wl_min, wl_max, output_excel,
                   time1, time2, time3, interval1, interval2, interval3):
    # 蒐集檔案與時間戳
    files = glob.glob(os.path.join(data_dir, '*.dat'))
    ts_map = {extract_timestamp_ms(fp): fp for fp in files if extract_timestamp_ms(fp) is not None}
    all_ms = sorted(ts_map.keys())

    if not all_ms:
        raise ValueError("資料夾內找不到符合規則的 .dat 檔（檔名需含數字時間戳）。")

    # 產生欲比對的時間點
    desired_secs = generate_time_points(start_s, time1, time2, time3, interval1, interval2, interval3)
    desired_ms = [int(round(s * 1000)) for s in desired_secs]

    spectra_list = []
    actual_secs = []
    max_wavelengths = []
    max_reflectances = []
    wavelengths = None

    for sec, dms in zip(desired_secs, desired_ms):
        # 找到第一個 >= dms 的檔
        best = next((ts for ts in all_ms if ts >= dms), None)
        if best is None:
            print(f"警告：找不到 >= {dms} ms ({sec} 秒) 的檔案，跳過。")
            continue

        wl, refl = read_spectrum(ts_map[best], wl_min, wl_max)

        if wl.size == 0:
            print(f"警告：{ts_map[best]} 在設定的波長範圍內沒有資料，跳過。")
            continue

        # 波長一致性檢查
        if wavelengths is None:
            wavelengths = wl
        elif not np.array_equal(wavelengths, wl):
            raise ValueError(f"檔案 {ts_map[best]} 的 wavelength 與先前不一致！")

        # 計算最大反射率與對應波長
        idx = int(np.nanargmax(refl))
        max_reflectances.append(float(refl[idx]))
        max_wavelengths.append(float(wl[idx]))

        spectra_list.append(refl)
        actual_secs.append(sec)

    if wavelengths is None or not spectra_list:
        raise ValueError("沒有可用資料，請檢查資料夾與檔名/波長範圍設定。")

    # 光譜矩陣
    df_spectra = pd.DataFrame(
        np.array(spectra_list).T,
        index=wavelengths,
        columns=[str(s) for s in actual_secs]
    )
    df_spectra.index.name = 'Wavelength (nm)'

    # 最大值表
    df_max = pd.DataFrame({
        '時間 (秒)': actual_secs,
        '最大反射率 (%)': max_reflectances,
        '最大反射波長 (nm)': max_wavelengths
    })

    # 輸出 Excel（需 openpyxl）
    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='w') as writer:
        df_spectra.to_excel(writer, sheet_name='Spectra')
        df_max.to_excel(writer,   sheet_name='MaxReflectance', index=False)

    print(f"已輸出 → {output_excel}")

# ---------- UI callbacks ----------
def loadFile():
    path = filedialog.askdirectory(title="請選擇資料夾")
    if path:
        loadFile_en.delete(0, tk.END)
        loadFile_en.insert(0, path)

def process_data():
    data_dir = loadFile_en.get().strip()
    if not os.path.isdir(data_dir):
        messagebox.showerror("錯誤", "請先選擇有效的資料夾！")
        return

    # 讀取輸入值
    try:
        wl_min    = float(entry4.get().strip())
        wl_max    = float(entry5.get().strip())
        start_s   = float(entry_start.get().strip())
        time1     = float(entry1.get().strip())
        time2     = float(entry2.get().strip())
        time3     = float(entry3.get().strip())
        interval1 = float(entry_time1.get().strip())
        interval2 = float(entry_time2.get().strip())
        interval3 = float(entry_time3.get().strip())
    except ValueError:
        messagebox.showerror("錯誤", "請在時間與波長欄位輸入有效數字！")
        return

    output_excel = os.path.join(data_dir, 'summary.xlsx')

    try:
        process_folder(
            data_dir, start_s, wl_min, wl_max, output_excel,
            time1, time2, time3, interval1, interval2, interval3
        )
    except Exception as e:
        messagebox.showerror("處理失敗", str(e))
    else:
        messagebox.showinfo("完成", f"已將結果儲存到\n{output_excel}")
        try:
            os.startfile(output_excel)  # Windows 開啟檔案的較穩做法
        except Exception:
            pass


#UI
win = tk.Tk()
win.title('Snake')
win.geometry('380x450')  
win.resizable(False, False)
icon_path = resource_path('icon.ico')
win.iconbitmap(icon_path)

# Label 區
tk.Label(win, text="請選取資料夾",       bg="grey", fg="white", height=1).place(x=0,   y=0)
tk.Label(win, text="時間範圍1 (s)",      fg="black", height=1).place(x=0,   y=50)
tk.Label(win, text="時間範圍2 (s)",      fg="black", height=1).place(x=0,   y=90)
tk.Label(win, text="時間範圍3 (s)",      fg="black",  height=1).place(x=0,   y=130)
tk.Label(win, text="波長範圍 (nm)",      fg="black",height=1).place(x=0,   y=190)
tk.Label(win, text="起始時間 (s)",       fg="black",height=1).place(x=0,   y=230)
tk.Label(win, text="~",       fg="black",height=1).place(x=215,   y=190)
# Entry 區
loadFile_en  = tk.Entry(win, width=35)
loadFile_en.place(x=80,  y=0)
entry1= tk.Entry(win, width=10)
entry1.place(x=120, y=50)
entry1.insert(0, "10")
entry2= tk.Entry(win, width=10)
entry2.place(x=120, y=90)
entry2.insert(0, "60")
entry3= tk.Entry(win, width=10)
entry3.place(x=120, y=130)
entry3.insert(0, "600")
entry4= tk.Entry(win, width=10)
entry4.place(x=120, y=190)
entry4.insert(0, "400")
entry5= tk.Entry(win, width=10)
entry5.place(x=250, y=190)
entry5.insert(0, "900")
entry_start  = tk.Entry(win, width=10)
entry_start.place(x=120, y=230)
entry_start.insert(0, "10")
entry_time1  = tk.Entry(win, width=10)
entry_time1.place(x=250, y=50)
entry_time1.insert(0, "1")
entry_time2  = tk.Entry(win, width=10)
entry_time2.place(x=250, y=90)
entry_time2.insert(0, "10")
entry_time3  = tk.Entry(win, width=10)
entry_time3.place(x=250, y=130)
entry_time3.insert(0, "30")

# Button 區
def infoboxOwen():
    messagebox.showinfo("Snake", "Program for In-situ Spectroscopy Analysis\n\nDesigned by Owen, NSYSU") #define Design info

tk.Button(win, text="...",        height=1, command=loadFile).place(x=355, y=0)
tk.Button(win, text="開始分析",    height=1, command=process_data).place(anchor=tk.CENTER, x=180, y=270)  # 調整 Y 座標
tk.Button(win, text="?",    height=1, command=infoboxOwen).place(x=5, y=420)

win.mainloop()

