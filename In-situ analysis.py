#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import glob
import numpy as np
import pandas as pd

# === 以下是原本的處理函式 ===

def extract_timestamp_ms(filepath):
    name = os.path.splitext(os.path.basename(filepath))[0]
    m = re.search(r'(\d+)', name)
    return int(m.group(1)) if m else None

def generate_time_points(start_s):
    offsets = []
    offsets += list(range(0, 11, 1))      # 0~10秒
    offsets += list(range(20, 61, 10))    # 20~60秒
    offsets += list(range(90, 601, 30))   # 90~600秒
    return [start_s + off for off in offsets]

def read_spectrum(filepath, wl_min, wl_max):
    df = pd.read_csv(filepath, sep='\t', skiprows=1, names=['PIXEL','WL','R'])
    mask = (df['WL'] >= wl_min) & (df['WL'] <= wl_max)
    return df.loc[mask, 'WL'].values, df.loc[mask, 'R'].values

def process_folder(data_dir, start_s, wl_min, wl_max, output_excel):
    files = glob.glob(os.path.join(data_dir, '*.dat'))
    ts_map = {extract_timestamp_ms(fp): fp for fp in files if extract_timestamp_ms(fp) is not None}
    all_ms = sorted(ts_map.keys())

    desired_secs = generate_time_points(start_s)
    desired_ms = [int(s * 1000) for s in desired_secs]

    spectra_list = []
    actual_secs = []
    max_wavelengths = []
    max_reflectances = []
    wavelengths = None

    for sec, dms in zip(desired_secs, desired_ms):
        best = next((ts for ts in all_ms if ts >= dms), None)
        wl, refl = read_spectrum(ts_map[best], wl_min, wl_max)

        # --- 計算最大反射率值及對應波長 ---
        idx = np.nanargmax(refl)
        max_reflectances.append(float(refl[idx]))
        max_wavelengths.append(float(wl[idx]))
        if best is None:
            print(f"警告：找不到 >= {dms} ms ({sec} 秒) 的檔案，跳過。")
            continue
        wl, refl = read_spectrum(ts_map[best], wl_min, wl_max)
        if wavelengths is None:
            wavelengths = wl
        elif not np.array_equal(wavelengths, wl):
            raise ValueError(f"檔案 {ts_map[best]} 的 wavelength 與先前不一致！")
        spectra_list.append(refl)
        actual_secs.append(sec)

    if wavelengths is None or not spectra_list:
        raise ValueError("沒有可用資料，請檢查資料夾與檔名格式。")

    df_spectra = pd.DataFrame(
        np.array(spectra_list).T,
        index=wavelengths,
        columns=[str(s) for s in actual_secs]
    )
    df_spectra.index.name = 'Wavelength (nm)'

    df_max = pd.DataFrame({
        '時間 (秒)': actual_secs,
        '最大反射率 (%)': max_reflectances,
        '最大反射波長(%)': max_wavelengths

    })

    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='w') as writer:
        df_spectra.to_excel(writer, sheet_name='Spectra')
        df_max.to_excel(writer,   sheet_name='MaxReflectance', index=False)

    print(f"已輸出 → {output_excel}")

# === 以下是 GUI 部分 ===

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

    # 讀取並驗證三個輸入框的值
    try:
        wl_min    = float(entry1.get().strip())
        wl_max    = float(entry2.get().strip())
        start_s   = float(entry_start.get().strip())  # ← 新增：讀取起始秒數
    except ValueError:
        messagebox.showerror("錯誤", "請在波長與起始秒數欄位輸入有效數字！")
        return

    output_excel = os.path.join(data_dir, 'summary.xlsx')

    try:
        process_folder(data_dir, start_s, wl_min, wl_max, output_excel)
    except Exception as e:
        messagebox.showerror("處理失敗", str(e))
    else:
        messagebox.showinfo("完成", f"已將結果儲存到\n{output_excel}")
        os.system(r'start excel.exe ' + output_excel)

# 建立主視窗
win = tk.Tk()
win.title('In-situ Spectra')
win.geometry('380x450')      # 改高一點以容納新欄位
win.resizable(False, False)
win.iconbitmap('C:/Users/user/iCloudDrive/NSYSU/tool/In-situ Spectra with GUI/ICON.jpg')

# Label 區
tk.Label(win, text="請選取資料夾",            bg="grey", fg="white", height=1).place(x=0,   y=0)
tk.Label(win, text="請輸入開始波長 (nm)",      fg="blue", height=1).place(x=0,   y=50)
tk.Label(win, text="請輸入結束波長 (nm)",      fg="red",  height=1).place(x=0,   y=120)
tk.Label(win, text="請輸入起始時間 (秒)",      fg="green",height=1).place(x=0,   y=190)  # ← 新增

# Entry 區
loadFile_en  = tk.Entry(win, width=40)
loadFile_en.place(x=70,  y=0)
entry1= tk.Entry(win, width=10)
entry1.place(x=120, y=50)
entry1.insert(0, "400")
entry2= tk.Entry(win, width=10)
entry2.place(x=120, y=120)
entry2.insert(0, "900")
entry_start  = tk.Entry(win, width=10)           # ← 新增
entry_start.place(x=120, y=190)
entry_start.insert(0, "10")  # 可以預設 5 秒

# Button 區
tk.Button(win, text="...",        height=1, command=loadFile).place(x=355, y=0)
tk.Button(win, text="開始分析",    height=1, command=process_data)\
  .place(anchor=tk.CENTER, x=180, y=270)  # 調整 Y 座標

win.mainloop()
