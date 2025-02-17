import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from datetime import datetime
import random
import calendar

class SchedulingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("林鼎鈞無償加班系統")
        self.employees = {}
        self.unavailable_dates = {}
        self.preassigned_shifts = {}
        self.exclusions = []
        self.shift_counts = {}

        # 選擇年月
        self.label_year = ttk.Label(root, text="選擇年份：")
        self.label_year.grid(row=0, column=0, padx=5, pady=5)
        current_year = datetime.today().year
        year_options = [str(year) for year in range(current_year, current_year + 30)]
        self.year_var = tk.StringVar(value=str(current_year))  # 設定預設值為當前年份
        self.entry_year = ttk.Combobox(root, textvariable=self.year_var, values=year_options, width=4, state="readonly")
        self.entry_year.grid(row=0, column=1, padx=5, pady=5)

        self.label_month = ttk.Label(root, text="選擇月份：")
        self.label_month.grid(row=0, column=2, padx=5, pady=5)
        current_month = datetime.today().month
        month_options = [str(month) for month in range(1, 13)]
        self.month_var = tk.StringVar(value=str(current_month))
        self.entry_month = ttk.Combobox(root, textvariable=self.month_var, values=month_options, width=2, state="readonly")
        self.entry_month.grid(row=0, column=3, padx=5, pady=5)

        # 員工 & 不可排班日 & 班次需求
        self.label_name = ttk.Label(root, text="員工姓名：")
        self.label_name.grid(row=1, column=0, padx=5, pady=5)
        self.entry_name = ttk.Entry(root, width=10)
        self.entry_name.grid(row=1, column=1, padx=5, pady=5)

        self.label_unavailable = ttk.Label(root, text="不可排班日：")
        self.label_unavailable.grid(row=1, column=2, padx=5, pady=5)
        self.entry_unavailable = ttk.Entry(root)
        self.entry_unavailable.grid(row=1, column=3, padx=5, pady=5)

        self.label_shift = ttk.Label(root, text="一線班數：")
        self.label_shift.grid(row=2, column=0, padx=5, pady=5)
        self.entry_shift_primary = ttk.Entry(root, width=5)
        self.entry_shift_primary.grid(row=2, column=1, padx=5, pady=5)

        self.label_shift = ttk.Label(root, text="二線班數：")
        self.label_shift.grid(row=2, column=2, padx=5, pady=5)
        self.entry_shift_secondary = ttk.Entry(root, width=5)
        self.entry_shift_secondary.grid(row=2, column=3, padx=5, pady=5)

        self.btn_add = ttk.Button(root, text="新增員工", command=self.add_employee)
        self.btn_add.grid(row=3, column=3, padx=5, pady=5)

        self.tree = ttk.Treeview(root, columns=("員工", "不可排班日", "一線", "二線"), show="headings")
        self.tree.heading("員工", text="員工")
        self.tree.heading("不可排班日", text="不可排班日")
        self.tree.heading("一線", text="一線班數")
        self.tree.heading("二線", text="二線班數")
        self.tree.grid(row=4, column=0, columnspan=4, padx=5, pady=5)

        # 預先排班
        self.label_year = ttk.Label(root, text="預先排班人員：")
        self.label_year.grid(row=5, column=0, padx=5, pady=5)
        self.combo_employee = ttk.Combobox(root, values=list(self.employees.keys()), state="readonly", width=10)
        self.combo_employee.grid(row=5, column=1, padx=5, pady=5)

        self.label_year = ttk.Label(root, text="日期：")
        self.label_year.grid(row=5, column=2, padx=5, pady=5)
        self.date_options = []
        self.combo_date = ttk.Combobox(root, values=self.date_options, state="readonly", width=2)
        self.combo_date.grid(row=5, column=3, padx=5, pady=5)
        self.entry_year.bind("<<ComboboxSelected>>", self.update_dates)
        self.entry_month.bind("<<ComboboxSelected>>", self.update_dates)
        
        self.label_year = ttk.Label(root, text="一線 / 二線：")
        self.label_year.grid(row=6, column=0, padx=5, pady=5)
        self.combo_shift_type = ttk.Combobox(root, values=["一線", "二線"], state="readonly", width=10)
        self.combo_shift_type.grid(row=6, column=1, padx=5, pady=5)

        self.btn_preassign = ttk.Button(root, text="預先排班", command=self.preassign_shift)
        self.btn_preassign.grid(row=6, column=3, columnspan=2, padx=5, pady=5)

        self.update_dates()

    def update_dates(self, event=None):
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        _, num_days = calendar.monthrange(year, month)
        self.date_options = [str(day) for day in range(1, num_days + 1)]
        self.combo_date["values"] = self.date_options
        self.combo_date.set('')

    def add_employee(self):
        name = self.entry_name.get().strip()
        if name and name not in self.employees:
            unavailable_dates_str = self.entry_unavailable.get().strip()
            unavailable_dates_list = [date.strip() for date in unavailable_dates_str.split(',') if date.strip().isdigit()]
            unavailable_dates_list = sorted(set(map(int, unavailable_dates_list)))
            primary_shifts = self.entry_shift_primary.get().strip()
            secondary_shifts = self.entry_shift_secondary.get().strip()
            if not primary_shifts.isdigit() or not secondary_shifts.isdigit():
                messagebox.showwarning("錯誤", "請輸入有效的班次數量")
                return
            self.employees[name] = []
            self.unavailable_dates[name] = unavailable_dates_list
            self.shift_counts[name] = {"一線": int(primary_shifts), "二線": int(secondary_shifts)}
            self.tree.insert("", tk.END, iid=name, values=(name, ", ".join(map(str, unavailable_dates_list)), primary_shifts, secondary_shifts))
            self.combo_employee["values"] = list(self.employees.keys())
            self.entry_name.delete(0, tk.END)
            self.entry_unavailable.delete(0, tk.END)
            self.entry_shift_primary.delete(0, tk.END)
            self.entry_shift_secondary.delete(0, tk.END)
        else:
            messagebox.showwarning("錯誤", "請輸入有效的員工名稱或避免重複")

    def preassign_shift(self):
        # 確認所有欄位都已選擇
        selected_employee = self.combo_employee.get()
        selected_date = self.combo_date.get()
        selected_shift = self.combo_shift_type.get()

        if not selected_employee or not selected_date or not selected_shift:
            messagebox.showwarning("錯誤", "請選擇所有欄位！")
            return

        # 確認該員工的不可排班日是否包含選擇的日期
        if selected_date in self.unavailable_dates.get(selected_employee, []):
            messagebox.showwarning("錯誤", f"{selected_employee} 在 {selected_date} 不可排班！")
            return

        # 如果所有條件都符合，將該員工的排班記錄新增到 `preassigned_shifts` 中
        if selected_employee not in self.preassigned_shifts:
            self.preassigned_shifts[selected_employee] = {}

        self.preassigned_shifts[selected_employee][selected_date] = selected_shift
        messagebox.showinfo("成功", f"{selected_employee} 已經在 {selected_date} 預排為 {selected_shift}！")

'''
    def add_unavailable_date(self):
        name = self.combo_employee.get().strip()
        dates = self.entry_unavailable.get().strip()

        if not name:
            messagebox.showwarning("錯誤", "請先選擇員工")
            return

        if name not in self.unavailable_dates:
            self.unavailable_dates[name] = []

        if dates:
            try:
                days = [int(day) for day in dates.split(',') if day.strip().isdigit()]
                self.unavailable_dates[name].extend(days)
                self.unavailable_dates[name] = sorted(set(self.unavailable_dates[name]))  # 移除重複並排序
                self.tree.item(name, values=(
                    name,
                    self.shift_counts[name]["一線"],
                    self.shift_counts[name]["二線"],
                    ', '.join(map(str, self.unavailable_dates[name]))
                ))
                self.entry_unavailable.delete(0, tk.END)
            except ValueError:
                messagebox.showwarning("錯誤", "請輸入有效的數字日期（用逗號分隔）")
        else:
            messagebox.showwarning("錯誤", "請輸入不可排班日")
    
    def add_preassigned_shift(self):
        name = self.combo_employee.get()
        day = self.entry_day.get().strip()
        shift_type = self.combo_shift.get()
        
        if not name or not day.isdigit() or not shift_type:
            messagebox.showwarning("錯誤", "請選擇員工並輸入有效日期和班次")
            return
        
        day = int(day)
        if name not in self.preassigned_shifts:
            self.preassigned_shifts[name] = {}
        self.preassigned_shifts[name][day] = shift_type
        
        messagebox.showinfo("成功", f"{name} 已被指定 {day} 號值 {shift_type} 班")
        self.entry_day.delete(0, tk.END)
    
    def generate_schedule(self):
        messagebox.showinfo("提示", "排班功能尚未實作！")
'''



if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulingApp(root)
    root.mainloop()