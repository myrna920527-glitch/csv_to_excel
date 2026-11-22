import os
import re
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.utils import get_column_letter


# 存放清單狀態的簡單全域變數（好維護、不用寫類別）
csv_files = []       # 完整檔案路徑清單
check_vars = []      # 每一列對應的勾選布林變數 (tk.BooleanVar)


def get_app_dir():
    """取得執行檔案（或 .py 腳本）所在的目錄"""
    if getattr(sys, "frozen", False):
        # 打包成 .exe 後的執行檔所在目錄
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def sanitize_sheet_name(name, used_names):
    """清理並確保符合 Excel 工作表（Sheet）命名規則"""
    # Excel 分頁名稱不可包含 \ / ? * : [ ]
    cleaned = re.sub(r'[\\/?*:\[\]]', "_", name).strip()
    if not cleaned:
        cleaned = "Sheet"

    # Excel 分頁名稱最長 31 個字元
    cleaned = cleaned[:31]

    # 若遇到重複名稱，加上編號避免報錯
    candidate = cleaned
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = f"{cleaned[:31 - len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate)
    return candidate


def read_csv_safely(csv_path):
    """嘗試常見編碼讀取 CSV，並將所有欄位讀為字串（避免前導零或店號被吃掉）"""
    encodings = ["utf-8-sig", "utf-8", "cp950"]
    last_error = None

    for enc in encodings:
        try:
            return pd.read_csv(csv_path, encoding=enc, dtype=str).fillna("")
        except UnicodeDecodeError as e:
            last_error = e

    raise last_error


def auto_adjust_column_width(worksheet):
    """根據內容長度自動調整 Excel 欄位寬度（中文字算 2、英數算 1）"""
    for col_idx, column_cells in enumerate(worksheet.columns, start=1):
        max_len = 0
        for cell in column_cells:
            val = str(cell.value) if cell.value is not None else ""
            # 計算顯示寬度：ASCII 字元算 1，非 ASCII（如中文字）算 2
            display_len = sum(2 if ord(ch) > 127 else 1 for ch in val)
            if display_len > max_len:
                max_len = display_len

        col_letter = get_column_letter(col_idx)
        # 留一點緩衝空間（+4），並設定最大上限避免超長文字把欄位撐爆
        worksheet.column_dimensions[col_letter].width = min(max_len + 4, 60)


def convert_separate_files(selected_files, output_dir):
    """獨立模式：每個勾選的 CSV 各轉成一個 .xlsx"""
    success_count = 0
    error_messages = []

    for csv_path in selected_files:
        try:
            df = read_csv_safely(csv_path)
            base_name = os.path.splitext(os.path.basename(csv_path))[0]
            used_names = set()
            sheet_name = sanitize_sheet_name(base_name, used_names)
            xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")

            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
                ws = writer.sheets[sheet_name]
                auto_adjust_column_width(ws)

            success_count += 1
        except Exception as e:
            error_messages.append(f"{os.path.basename(csv_path)}：{e}")

    return success_count, error_messages, output_dir


def convert_merged_file(selected_files, output_dir):
    """合併模式：將勾選的 CSV 合併進同一個 .xlsx，每個檔案一個分頁"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    merged_xlsx_path = os.path.join(output_dir, f"merged_{timestamp}.xlsx")

    used_sheet_names = set()
    success_count = 0
    error_messages = []

    # 先讀取所有勾選的 CSV，避免寫入一半才發現某檔壞掉
    loaded_items = []
    for csv_path in selected_files:
        try:
            df = read_csv_safely(csv_path)
            base_name = os.path.splitext(os.path.basename(csv_path))[0]
            sheet_name = sanitize_sheet_name(base_name, used_sheet_names)
            loaded_items.append((csv_path, df, sheet_name))
        except Exception as e:
            error_messages.append(f"{os.path.basename(csv_path)}：{e}")

    if not loaded_items:
        return 0, error_messages, merged_xlsx_path

    with pd.ExcelWriter(merged_xlsx_path, engine="openpyxl") as writer:
        for _, df, sheet_name in loaded_items:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            auto_adjust_column_width(ws)
            success_count += 1

    return success_count, error_messages, merged_xlsx_path


def clear_checkbox_list():
    """清空目前的檔案清單與勾選方塊"""
    for widget in list_frame.winfo_children():
        widget.destroy()
    csv_files.clear()
    check_vars.clear()


def select_directory():
    """選擇 CSV 目錄並列出所有 .csv 檔案"""
    folder_path = filedialog.askdirectory(title="請選擇包含 CSV 的目錄")
    if not folder_path:
        return

    dir_var.set(folder_path)
    clear_checkbox_list()

    found_files = [
        os.path.join(folder_path, f)
        for f in sorted(os.listdir(folder_path))
        if f.lower().endswith(".csv") and os.path.isfile(os.path.join(folder_path, f))
    ]

    if not found_files:
        tk.Label(list_frame, text="此目錄下找不到任何 .csv 檔案", fg="gray").pack(anchor="w", padx=6, pady=6)
        return

    for file_path in found_files:
        var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(list_frame, text=os.path.basename(file_path), variable=var)
        chk.pack(anchor="w", padx=6, pady=2)

        csv_files.append(file_path)
        check_vars.append(var)


def set_all_checks(state):
    """全選或全取消勾選"""
    for var in check_vars:
        var.set(state)


def start_convert():
    """取得被勾選的 CSV，依所選模式轉成 Excel 並輸出到執行檔旁的 output 目錄"""
    selected_files = [
        path for path, var in zip(csv_files, check_vars) if var.get()
    ]

    if not selected_files:
        messagebox.showwarning("提醒", "請先選擇目錄，並至少勾選一個 CSV 檔案！")
        return

    output_dir = os.path.join(get_app_dir(), "output")
    os.makedirs(output_dir, exist_ok=True)

    mode = mode_var.get()
    if mode == "merge":
        success_count, error_messages, target_path = convert_merged_file(selected_files, output_dir)
        result_msg = f"合併完成！共寫入 {success_count} 個分頁。\n檔案位置：{target_path}"
    else:
        success_count, error_messages, target_path = convert_separate_files(selected_files, output_dir)
        result_msg = f"轉換完成！共成功 {success_count} 個檔案。\n輸出目錄：{target_path}"

    if error_messages:
        result_msg += "\n\n以下檔案發生問題：\n" + "\n".join(error_messages)
        messagebox.showwarning("部分完成", result_msg)
    elif success_count > 0:
        messagebox.showinfo("完成", result_msg)
    else:
        messagebox.showerror("失敗", result_msg)


# ================= 建立視窗介面 =================
root = tk.Tk()
root.title("CSV 轉 Excel 小工具")
root.geometry("540x470")

dir_var = tk.StringVar(value="尚未選擇目錄")
mode_var = tk.StringVar(value="merge")  # 預設為合併模式

# 1. 目錄選擇列
top_frame = tk.Frame(root, padx=10, pady=10)
top_frame.pack(fill="x")

tk.Button(top_frame, text="選擇 CSV 目錄", command=select_directory).pack(side="left")
tk.Label(top_frame, textvariable=dir_var, fg="#333333", padx=8).pack(side="left", fill="x")

# 2. 轉換模式與全選列
option_frame = tk.Frame(root, padx=10, pady=4)
option_frame.pack(fill="x")

tk.Label(option_frame, text="輸出模式：").pack(side="left")
tk.Radiobutton(
    option_frame,
    text="合併成單一 Excel（各 CSV 為分頁）",
    variable=mode_var,
    value="merge"
).pack(side="left", padx=(0, 10))
tk.Radiobutton(
    option_frame,
    text="各自獨立轉檔",
    variable=mode_var,
    value="separate"
).pack(side="left")

action_bar = tk.Frame(root, padx=10, pady=4)
action_bar.pack(fill="x")

tk.Button(action_bar, text="全選", command=lambda: set_all_checks(True)).pack(side="left", padx=(0, 5))
tk.Button(action_bar, text="全不選", command=lambda: set_all_checks(False)).pack(side="left")

# 3. 可捲動的勾選清單區
container = tk.LabelFrame(root, text="CSV 檔案清單（勾選要轉換的檔案）", padx=8, pady=8)
container.pack(fill="both", expand=True, padx=10, pady=8)

canvas = tk.Canvas(container, highlightthickness=0)
scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
list_frame = tk.Frame(canvas)

list_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=list_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# 4. 底部一鍵轉換按鈕
bottom_frame = tk.Frame(root, padx=10, pady=10)
bottom_frame.pack(fill="x")

tk.Button(
    bottom_frame,
    text="一鍵轉換成 Excel",
    command=start_convert,
    height=2,
    bg="#2E7D32",
    fg="white",
    font=("微軟正黑體", 11, "bold")
).pack(fill="x")

root.mainloop()