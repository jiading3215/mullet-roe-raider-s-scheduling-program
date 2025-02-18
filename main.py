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

        # 員工基本資料表格
        self.employee_tree = ttk.Treeview(root, columns=("員工", "不可排班日", "一線", "二線"), show="headings")
        self.employee_tree.heading("員工", text="員工")
        self.employee_tree.heading("不可排班日", text="不可排班日")
        self.employee_tree.heading("一線", text="一線班數")
        self.employee_tree.heading("二線", text="二線班數")
        self.employee_tree.grid(row=4, column=0, columnspan=4, padx=5, pady=5)

        # 預先排班
        self.label_employee = ttk.Label(root, text="預先排班人員：")
        self.label_employee.grid(row=5, column=0, padx=5, pady=5)
        self.combo_employee = ttk.Combobox(root, values=list(self.employees.keys()), state="readonly", width=10)
        self.combo_employee.grid(row=5, column=1, padx=5, pady=5)

        self.label_date = ttk.Label(root, text="日期：")
        self.label_date.grid(row=5, column=2, padx=5, pady=5)
        self.date_options = []
        self.combo_date = ttk.Combobox(root, values=self.date_options, state="readonly", width=2)
        self.combo_date.grid(row=5, column=3, padx=5, pady=5)
        self.entry_year.bind("<<ComboboxSelected>>", self.update_dates)
        self.entry_month.bind("<<ComboboxSelected>>", self.update_dates)
        
        self.label_shift_type = ttk.Label(root, text="一線 / 二線：")
        self.label_shift_type.grid(row=6, column=0, padx=5, pady=5)
        self.combo_shift_type = ttk.Combobox(root, values=["一線", "二線"], state="readonly", width=10)
        self.combo_shift_type.grid(row=6, column=1, padx=5, pady=5)

        self.btn_preassign = ttk.Button(root, text="預先排班", command=self.preassign_shift)
        self.btn_preassign.grid(row=6, column=3, columnspan=2, padx=5, pady=5)

        self.update_dates()

        # 預排班次結果表格
        self.shift_tree = ttk.Treeview(root, columns=("員工", "日期", "班次"), show="headings")
        self.shift_tree.heading("員工", text="員工")
        self.shift_tree.heading("日期", text="日期")
        self.shift_tree.heading("班次", text="班次")
        self.shift_tree.grid(row=7, column=0, columnspan=4, padx=5, pady=5)

        # 禁排組合
        self.label_exclusion = ttk.Label(root, text="禁排員工組合：")
        self.label_exclusion.grid(row=8, column=0, padx=5, pady=5)
        self.combo_exclusion_employee_1 = ttk.Combobox(root, values=list(self.employees.keys()), state="readonly", width=10)
        self.combo_exclusion_employee_1.grid(row=8, column=1, padx=5, pady=5)
        self.combo_exclusion_employee_2 = ttk.Combobox(root, values=list(self.employees.keys()), state="readonly", width=10)
        self.combo_exclusion_employee_2.grid(row=8, column=2, padx=5, pady=5)
        self.btn_add_exclusion = ttk.Button(root, text="新增禁排組合", command=self.add_exclusion)
        self.btn_add_exclusion.grid(row=8, column=3, padx=5, pady=5)

        self.exclusion_tree = ttk.Treeview(root, columns=("員工1", "員工2"), show="headings")
        self.exclusion_tree.heading("員工1", text="員工1")
        self.exclusion_tree.heading("員工2", text="員工2")
        self.exclusion_tree.grid(row=9, column=0, columnspan=4, padx=5, pady=5)

        self.btn_generate_excel = ttk.Button(root, text="生成班表", command=self.generate_excel_schedule)
        self.btn_generate_excel.grid(row=10, column=0, columnspan=4, padx=5, pady=5)

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
            self.employee_tree.insert("", tk.END, iid=name, values=(name, ", ".join(map(str, unavailable_dates_list)), primary_shifts, secondary_shifts))
            self.combo_employee["values"] = list(self.employees.keys())
            self.combo_exclusion_employee_1["values"] = list(self.employees.keys())  # 更新禁排組合1
            self.combo_exclusion_employee_2["values"] = list(self.employees.keys())  # 更新禁排組合2
            self.entry_name.delete(0, tk.END)
            self.entry_unavailable.delete(0, tk.END)
            self.entry_shift_primary.delete(0, tk.END)
            self.entry_shift_secondary.delete(0, tk.END)
        else:
            messagebox.showwarning("錯誤", "請輸入有效的員工名稱或避免重複")

    def preassign_shift(self):
        employee = self.combo_employee.get()
        date = self.combo_date.get()
        shift_type = self.combo_shift_type.get()
        if not employee or not date or not shift_type:
            messagebox.showwarning("錯誤", "請選擇員工、日期和班次")
            return
        # 檢查員工是否有在選擇的日期有不可排班日
        if int(date) in self.unavailable_dates.get(employee, []):
            messagebox.showwarning("錯誤", f"{employee} 在 {date} 不可排班")
            return
        # 如果員工尚未有預排班次，則創建一個新的列表
        if employee not in self.preassigned_shifts:
            self.preassigned_shifts[employee] = []
        # 確保這個員工在這個日期和班次上還沒排過班
        if any(item['date'] == date and item['shift'] == shift_type for item in self.preassigned_shifts[employee]):
            messagebox.showwarning("錯誤", f"{employee} 已在 {date} 排過 {shift_type} 班")
            return
        # 將新的排班信息加入列表
        self.preassigned_shifts[employee].append({'date': date, 'shift': shift_type})
        messagebox.showinfo("成功", f"成功為 {employee} 預排 {date} 的 {shift_type} 班")
        self.update_preassigned_shifts_treeview()

    def update_preassigned_shifts_treeview(self):
        # 清空 Treeview
        for row in self.shift_tree.get_children():
            self.shift_tree.delete(row)
        # 將 preassigned_shifts 中的資料顯示在 Treeview
        for employee, shifts in self.preassigned_shifts.items():
            for shift in shifts:
                self.shift_tree.insert("", tk.END, values=(employee, shift['date'], shift['shift']))

    def add_exclusion(self):
        employee_1 = self.combo_exclusion_employee_1.get().strip()
        employee_2 = self.combo_exclusion_employee_2.get().strip()
        # 檢查是否兩個員工都選擇了
        if not employee_1 or not employee_2:
            messagebox.showwarning("錯誤", "請選擇兩個員工")
            return
        # 檢查是否選擇的是相同的員工
        if employee_1 == employee_2:
            messagebox.showwarning("錯誤", "兩個員工不能相同")
            return
        # 將禁排組合添加到列表中 (以小的名稱在前，較大的名稱在後，以避免重複)
        exclusion_combination = sorted([employee_1, employee_2])
        if exclusion_combination not in self.exclusions:
            self.exclusions.append(exclusion_combination)
            messagebox.showinfo("成功", f"成功新增禁排組合：{employee_1} 和 {employee_2}")
            self.update_exclusion_treeview()  # 更新禁排組合表格
        else:
            messagebox.showwarning("錯誤", "這個禁排組合已經存在")
        # 更新員工選單，以確保新員工加入時選單顯示最新員工列表
        self.combo_exclusion_employee_1["values"] = list(self.employees.keys())
        self.combo_exclusion_employee_2["values"] = list(self.employees.keys())

    def update_exclusion_treeview(self):
        # 清空 Treeview
        for row in self.exclusion_tree.get_children():
            self.exclusion_tree.delete(row)
        # 將禁排組合顯示在 Treeview
        for exclusion in self.exclusions:
            self.exclusion_tree.insert("", tk.END, values=(exclusion[0], exclusion[1]))

    def assign_shifts(self):
        # 1. 優先處理預先排班的員工
        for employee, shifts in self.preassigned_shifts.items():
            for shift in shifts:
                self.schedule_shift(employee, shift['date'], shift['shift'])
        
        # 2. 根據每位員工的需求排班
        for employee, shift_info in self.shift_counts.items():
            total_shifts_needed = shift_info["一線"] + shift_info["二線"]
            unavailable_dates = self.unavailable_dates.get(employee, [])
            
            # 根據總排班次數排班，間隔最大化
            shift_dates = self.schedule_shifts_for_employee(employee, total_shifts_needed, unavailable_dates)
            
            # 更新員工的排班資訊
            self.schedule[employee] = shift_dates

    def schedule_shifts_for_employee(self, employee, total_shifts_needed, unavailable_dates):
        scheduled_dates = []  # 儲存排班的日期
        all_dates = list(range(1, 31))  # 假設每個月有30天
        available_dates = [date for date in all_dates if date not in unavailable_dates]

        # 根據排班次數計算最大間隔
        gap = len(available_dates) // total_shifts_needed

        while total_shifts_needed > 0 and available_dates:
            # 找到最大間隔的可排日期
            possible_dates = [date for date in available_dates if not self.is_consecutive_shift(employee, date)]
            if not possible_dates:
                break

            # 根據計算的間隔挑選日期
            selected_date = available_dates[min(range(len(available_dates)), key=lambda i: abs(available_dates[i] - gap))]
            self.schedule_shift(employee, selected_date, "未指定")  # "未指定"代表班次為待定
            scheduled_dates.append(selected_date)
            available_dates.remove(selected_date)
            total_shifts_needed -= 1

        return scheduled_dates

    def schedule_shift(self, employee, date, shift_type):
        # 檢查禁排組合
        if any((employee, other) in self.exclusions or (other, employee) in self.exclusions for other in self.schedule):
            print(f"禁排組合：{employee} 在 {date} 不可與其他員工排班")
            return
        
        # 確認該日期和班次不重複
        if employee not in self.schedule:
            self.schedule[employee] = []
        self.schedule[employee].append({'date': date, 'shift': shift_type})

    def is_consecutive_shift(self, employee, date):
        # 檢查該員工是否在日期前後有排班
        for shift in self.schedule.get(employee, []):
            if abs(shift['date'] - date) == 1:  # 檢查是否為連續排班
                return True
        return False

    def print_schedule(self):
        # 輸出排班結果
        for employee, shifts in self.schedule.items():
            print(f"{employee} 排班：")
            for shift in shifts:
                print(f"  日期: {shift['date']}, 班次: {shift['shift']}")

    def generate_excel_schedule(self):
        # 獲取所選的年份和月份
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        
        # 根據年份和月份獲取該月的天數
        _, num_days = calendar.monthrange(year, month)
        
        # 創建一個空的 DataFrame 來存儲排班結果
        columns = ['日期', '一線班', '二線班']
        schedule_data = []

        # 遍歷每一天的排班
        for date in range(1, num_days + 1):  # 動態處理該月的天數
            one_line_shifts = []  # 儲存一線班的員工
            two_line_shifts = []  # 儲存二線班的員工

            # 查找每個員工是否在該日期有排班
            for employee, shifts in self.schedule.items():
                for shift in shifts:
                    if shift['date'] == date:
                        if shift['shift'] == "一線":
                            one_line_shifts.append(employee)
                        elif shift['shift'] == "二線":
                            two_line_shifts.append(employee)

            # 將該日期的排班添加到資料中
            schedule_data.append([date, ', '.join(one_line_shifts), ', '.join(two_line_shifts)])

        # 使用 pandas 創建 DataFrame
        df = pd.DataFrame(schedule_data, columns=columns)

        # 儲存為 Excel 文件
        file_name = f"排班表_{year}_{month}.xlsx"
        df.to_excel(file_name, index=False, engine='openpyxl')

        messagebox.showinfo("成功", f"排班表已生成：{file_name}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulingApp(root)
    root.mainloop()