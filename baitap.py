import tkinter as tk
from tkinter import messagebox
import csv
import pandas as pd
from datetime import datetime

# Tên file CSV và Excel
CSV_FILE = "employees.csv"
EXCEL_FILE = "employees_sorted.xlsx"

# Tạo file CSV nếu chưa tồn tại
def create_csv_if_not_exists():
    try:
        with open(CSV_FILE, 'x', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", 
                             "Số CMND", "Ngày cấp", "Nơi cấp", "Là khách hàng", "Là nhà cung cấp"])
    except FileExistsError:
        pass

# Lưu dữ liệu vào file CSV
def save_to_csv():
    data = [
        entry_ma.get(), entry_ten.get(), entry_don_vi.get(), entry_chuc_danh.get(),
        entry_ngay_sinh.get(), gender_var.get(), entry_cmnd.get(), entry_ngay_cap.get(), 
        entry_noi_cap.get(), customer_var.get(), supplier_var.get()
    ]
    if "" in data[:9]:  # Kiểm tra các ô nhập chính có dữ liệu không
        messagebox.showerror("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
        return

    # Ghi vào file CSV
    with open(CSV_FILE, 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(data)
    messagebox.showinfo("Thành công", "Dữ liệu đã được lưu!")
    clear_entries()

# Xóa các ô nhập liệu
def clear_entries():
    entry_ma.delete(0, tk.END)
    entry_ten.delete(0, tk.END)
    entry_don_vi.delete(0, tk.END)
    entry_chuc_danh.delete(0, tk.END)
    entry_ngay_sinh.delete(0, tk.END)
    gender_var.set("Nam")
    entry_cmnd.delete(0, tk.END)
    entry_ngay_cap.delete(0, tk.END)
    entry_noi_cap.delete(0, tk.END)
    customer_var.set(0)
    supplier_var.set(0)

# Tìm nhân viên có sinh nhật hôm nay
def find_birthdays_today():
    today = datetime.now().strftime("%d/%m")
    results = []
    with open(CSV_FILE, 'r') as file:
        reader = csv.reader(file)
        next(reader)
        for row in reader:
            if today in row[4]:
                results.append(row)
    if results:
        message = "Nhân viên có sinh nhật hôm nay:\n" + "\n".join([f"{r[1]} ({r[0]})" for r in results])
    else:
        message = "Không có nhân viên nào có sinh nhật hôm nay."
    messagebox.showinfo("Sinh nhật hôm nay", message)

# Xuất danh sách nhân viên ra file Excel
def export_to_excel():
    try:
        df = pd.read_csv(CSV_FILE)
        df['Ngày sinh'] = pd.to_datetime(df['Ngày sinh'], format='%d/%m/%Y', errors='coerce')
        df = df.sort_values(by='Ngày sinh', ascending=True)
        df.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Thành công", f"Dữ liệu đã được xuất ra file '{EXCEL_FILE}'!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Xuất file thất bại: {e}")

# Tạo giao diện người dùng
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")
root.geometry("900x500")

# Tiêu đề chính
title_label = tk.Label(root, text="Thông tin nhân viên", font=("Arial", 16, "bold"))
title_label.grid(row=0, column=0, columnspan=2, pady=10)

# Các biến lưu trữ dữ liệu
gender_var = tk.StringVar(value="Nam")
customer_var = tk.IntVar()
supplier_var = tk.IntVar()

# Các Label và Entry nhập liệu
tk.Label(root, text="Mã *:").grid(row=1, column=0, padx=5, pady=5)
entry_ma = tk.Entry(root)
entry_ma.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Tên *:").grid(row=1, column=2, padx=5, pady=5)
entry_ten = tk.Entry(root)
entry_ten.grid(row=1, column=3, padx=5, pady=5)

tk.Label(root, text="Đơn vị *:").grid(row=2, column=0, padx=5, pady=5)
entry_don_vi = tk.Entry(root)
entry_don_vi.grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Chức danh:").grid(row=2, column=2, padx=5, pady=5)
entry_chuc_danh = tk.Entry(root)
entry_chuc_danh.grid(row=2, column=3, padx=5, pady=5)

tk.Label(root, text="Ngày sinh *:").grid(row=3, column=0, padx=5, pady=5)
entry_ngay_sinh = tk.Entry(root)
entry_ngay_sinh.grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Giới tính:").grid(row=3, column=2, padx=5, pady=5)
tk.Radiobutton(root, text="Nam", variable=gender_var, value="Nam").grid(row=3, column=3, sticky="w")
tk.Radiobutton(root, text="Nữ", variable=gender_var, value="Nữ").grid(row=3, column=3, padx=50, sticky="w")

tk.Label(root, text="Số CMND:").grid(row=4, column=0, padx=5, pady=5)
entry_cmnd = tk.Entry(root)
entry_cmnd.grid(row=4, column=1, padx=5, pady=5)

tk.Label(root, text="Ngày cấp:").grid(row=4, column=2, padx=5, pady=5)
entry_ngay_cap = tk.Entry(root)
entry_ngay_cap.grid(row=4, column=3, padx=5, pady=5)

tk.Label(root, text="Nơi cấp:").grid(row=5, column=0, padx=5, pady=5)
entry_noi_cap = tk.Entry(root)
entry_noi_cap.grid(row=5, column=1, padx=5, pady=5)

# Checkbox "Là khách hàng" và "Là nhà cung cấp"
tk.Checkbutton(root, text="Là khách hàng", variable=customer_var).grid(row=0 , column=2, sticky="w", padx=5, pady=5)
tk.Checkbutton(root, text="Là nhà cung cấp", variable=supplier_var).grid(row=0 , column=3, sticky="w", padx=5, pady=5)

# Các nút chức năng
tk.Button(root, text="Lưu dữ liệu", command=save_to_csv).grid(row=6, column=0, pady=10)
tk.Button(root, text="Sinh nhật ngày hôm nay", command=find_birthdays_today).grid(row=6, column=1, pady=10)
tk.Button(root, text="Xuất toàn bộ danh sách", command=export_to_excel).grid(row=6, column=2, pady=10)
tk.Button(root, text="Thoát", command=root.quit).grid(row=6, column=3, pady=10)

create_csv_if_not_exists()
root.mainloop()

