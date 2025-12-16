import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv

df_global = None
version = '1.2.1'

# تابع پیدا کردن مسیر درست آیکون (مخصوص EXE)
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def browse_csv():
    global df_global
    file_path = filedialog.askopenfilename(filetypes=[("CSV or Excel Files", "*.csv *.xlsx *.xls")])
    if file_path:
        csv_entry.delete(0, tk.END)
        csv_entry.insert(0, file_path)

        try:
            if file_path.endswith('.csv'):
                delimiter = delimiter_entry.get() or ','
                remove_bad_rows = remove_bad_rows_var.get()

                with open(file_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    header = next(reader)
                    expected_columns_count = len(header)

                    if remove_bad_rows:
                        rows = [header]
                        bad_rows_count = 0
                        for i, row in enumerate(reader, start=2):
                            if len(row) == expected_columns_count:
                                rows.append(row)
                            else:
                                bad_rows_count += 1
                        if bad_rows_count > 0:
                            messagebox.showinfo("اطلاع", f"{bad_rows_count} سطر دارای تعداد ستون نامناسب حذف شد.")
                        df = pd.DataFrame(rows[1:], columns=rows[0])
                    else:
                        for i, row in enumerate(reader, start=2):
                            if len(row) != expected_columns_count:
                                messagebox.showerror(
                                    "خطا در ساختار فایل",
                                    f"❌ تعداد ستون‌ها در ردیف {i} با هدر برابر نیست.\n"
                                    f"انتظار می‌رفت {expected_columns_count} ستون باشد، اما {len(row)} ستون یافت شد."
                                )
                                return
                        df = pd.read_csv(file_path, encoding='utf-8', sep=delimiter)

            elif file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='openpyxl' if file_path.endswith('.xlsx') else None)
                delimiter_entry.delete(0, tk.END)
                delimiter_entry.insert(0, ",")

            else:
                messagebox.showerror("خطا", "فرمت فایل پشتیبانی نمی‌شود. فقط CSV یا Excel مجاز است.")
                return

            df.columns = df.columns.str.strip()

            if any(not col or pd.isna(col) for col in df.columns):
                messagebox.showerror("خطا", "❌ یکی از ستون‌ها بدون نام است یا مقدار آن نامعتبر است.")
                return

            df_global = df
            columns = df.columns.tolist()
            column_dropdown['values'] = columns
            if columns:
                column_dropdown.current(0)

        except Exception as e:
            messagebox.showerror("خطا", f"❌ خطا در خواندن فایل:\n{e}")

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)

def process_csv():
    global df_global
    file_path = csv_entry.get()
    output_base = output_entry.get()
    group_column = column_dropdown.get().strip()
    delimiter = delimiter_entry.get()

    if df_global is None:
        messagebox.showerror("خطا", "لطفاً فایل ورودی را انتخاب کنید.")
        return

    if not output_base or not group_column:
        messagebox.showerror("خطا", "لطفاً پوشه خروجی و ستون را مشخص کنید.")
        return

    if not delimiter:
        messagebox.showerror("خطا", "لطفاً کاراکتر جداکننده را وارد کنید.")
        return

    try:
        df = df_global
        if group_column not in df.columns:
            messagebox.showerror("خطا", f'ستون "{group_column}" در فایل وجود ندارد.')
            return

        unique_groups = df[group_column].dropna().unique()

        for group in unique_groups:
            group_df = df[df[group_column] == group]
            row_count = len(group_df)

            safe_group_name = str(group).strip().replace('/', '_').replace('\\', '_') \
                                 .replace(':', '_').replace('*', '_').replace('?', '_') \
                                 .replace('"', '_').replace('<', '_').replace('>', '_') \
                                 .replace('|', '_').replace('\n', '').replace('\r', '').replace(' ', '_')

            group_folder = os.path.join(output_base, safe_group_name)
            os.makedirs(group_folder, exist_ok=True)

            output_file = os.path.join(group_folder, f'{safe_group_name} ({row_count}).csv')

            group_df.to_csv(output_file, index=False, encoding='utf-8-sig', sep=delimiter)

            with open(output_file, 'a', encoding='utf-8') as f:
                f.write('\n# تولید شده توسط برنامه AtaFilter | نسخه: ' + version + '\n')

        messagebox.showinfo("موفق", "✅ همه فایل‌ها با موفقیت ذخیره شدند.")
    except Exception as e:
        messagebox.showerror("خطا", f"❌ خطا در پردازش:\n{e}")

# ساخت پنجره اصلی
root = tk.Tk()
root.title("برنامه AtaFilter  نسخه: " + version)
root.geometry("700x400")
root.resizable(False, False)

# ست کردن آیکون از مسیر درست
try:
    root.iconbitmap(resource_path("app.ico"))
except Exception as e:
    print("❌ خطا در بارگذاری آیکون:", e)

# کمک‌کننده‌های راست‌چین
def rtl_label(label):
    label.config(anchor='e', justify='right', font=("Tahoma", 11))

def rtl_entry(entry):
    entry.config(justify='right', font=("Tahoma", 11))

def rtl_combobox(cb):
    cb.config(justify='right', font=("Tahoma", 11))

# ردیف اول: فایل CSV
label1 = tk.Label(root, text=":مسیر فایل CSV یا Excel")
label1.grid(row=0, column=2, sticky="w", padx=5, pady=10)
rtl_label(label1)

csv_entry = tk.Entry(root, width=50)
csv_entry.grid(row=0, column=1, padx=5, sticky="e")
rtl_entry(csv_entry)
tk.Button(root, text="انتخاب", command=browse_csv).grid(row=0, column=0, padx=5)

# ردیف دوم: پوشه خروجی
label2 = tk.Label(root, text=":مسیر ذخیره خروجی")
label2.grid(row=1, column=2, sticky="w", padx=5, pady=10)
rtl_label(label2)

output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=5, sticky="e")
rtl_entry(output_entry)
tk.Button(root, text="انتخاب", command=browse_output_folder).grid(row=1, column=0, padx=5)

# ردیف سوم: انتخاب ستون
label3 = tk.Label(root, text=":ستون گروه‌بندی")
label3.grid(row=2, column=2, sticky="w", padx=5, pady=10)
rtl_label(label3)

column_dropdown = ttk.Combobox(root, state="readonly", width=47)
column_dropdown.grid(row=2, column=1, padx=5, sticky="e")
rtl_combobox(column_dropdown)

# ردیف چهارم: جداکننده
label4 = tk.Label(root, text=":کاراکتر جداکننده")
label4.grid(row=3, column=2, sticky="w", padx=5, pady=10)
rtl_label(label4)

delimiter_entry = tk.Entry(root, width=10)
delimiter_entry.grid(row=3, column=1, sticky="w", padx=5)
rtl_entry(delimiter_entry)
delimiter_entry.insert(0, ",")

# حذف سطرهای مشکل‌دار
remove_bad_rows_var = tk.BooleanVar()
remove_bad_rows_check = tk.Checkbutton(root, text="حذف سطرهای دارای خطا", variable=remove_bad_rows_var, font=("Tahoma", 11))
remove_bad_rows_check.grid(row=4, column=1, sticky="e", padx=5, pady=10)

# دکمه اجرا
run_button = tk.Button(root, text="✅ اجرا", command=process_csv, bg="green", fg="white", font=("Tahoma", 14))
run_button.grid(row=5, column=1, pady=30, sticky="e")

root.mainloop()


