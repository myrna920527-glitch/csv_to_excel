import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.utils import get_column_letter


# 存放清單狀態的簡單全域變數（好維護、不用寫類別）
csv_files = []       # 完整檔案路徑清單
check_vars = []      # 每一列對應的勾選布林變數 (tk.BooleanVar)


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


def convert_single_csv(csv_path, output_dir):
    """將單一 CSV 轉成同名的 XLSX，並自動調整欄寬"""
    df = read_csv_safely(csv_path)

    base_name = os.path.splitext(os.path.basename(csv_path))[0]
    xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="資料")
        ws = writer.sheets["資料"]
        auto_adjust_column_width(ws)

    return xlsx_path


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

    # 找出目錄下的 .csv 檔並排序
    found_files = [
        os.path.join(folder_path, f)
        for f in sorted(os.listdir(folder_path))
        if f.lower().endswith(".csv") and os.path.isfile(os.path.join(folder_path, f))
    ]

    if not found_files:
        tk.Label(list_frame, text="此目錄下找不到任何 .csv 檔案", fg="gray").pack(anchor="w", padx=6, pady=6)
        return

    for file_path in found_files:
        var = tk.BooleanVar(value=True)  # 預設全部勾選，方便直接轉
        chk = tk.Checkbutton(list_frame, text=os.path.basename(file_path), variable=var)
        chk.pack(anchor="w", padx=6, pady=2)

        csv_files.append(file_path)
        check_vars.append(var)


def set_all_checks(state):
    """全選或全取消勾選"""
    for var in check_vars:
        var.set(state)


def start_convert():
    """取得被勾選的 CSV，批次轉成 Excel"""
    selected_files = [
        path for path, var in zip(csv_files, check_vars) if var.get()
    ]

    if not selected_files:
        messagebox.showwarning("提醒", "請先選擇目錄，並至少勾選一個 CSV 檔案！")
        return

    # 預設輸出到原目錄底下的 excel_output 資料夾（也可改成讓使用者另選）
    base_dir = dir_var.get()
    output_dir = os.path.join(base_dir, "excel_output")
    os.makedirs(output_dir, exist_ok=True)

    success_count = 0
    error_messages = []

    for csv_path in selected_files:
        try:
            convert_single_csv(csv_path, output_dir)
            success_count += 1
        except Exception as e:
            error_messages.append(f"{os.path.basename(csv_path)}：{e}")

    result_msg = f"轉換完成！共成功 {success_count} 個檔案。\n輸出位置：{output_dir}"
    if error_messages:
        result_msg += "\n\n以下檔案轉換失敗：\n" + "\n".join(error_messages)
        messagebox.showwarning("部分完成", result_msg)
    else:
        messagebox.showinfo("完成", result_msg)


# ================= 建立視窗介面 =================
root = tk.Tk()
root.title("CSV 轉 Excel 小工具")
root.geometry("520x420")

dir_var = tk.StringVar(value="尚未選擇目錄")

# 1. 目錄選擇列
top_frame = tk.Frame(root, padx=10, pady=10)
top_frame.pack(fill="x")

tk.Button(top_frame, text="選擇 CSV 目錄", command=select_directory).pack(side="left")
tk.Label(top_frame, textvariable=dir_var, fg="#333333", padx=8).pack(side="left", fill="x")

# 2. 全選 / 全取消按鈕列
action_bar = tk.Frame(root, padx=10)
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