import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import calendar

class SchedulingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("林鼎鈞無償加班系統")
        self.employees = {}
        self.unavailable_dates = {}
        self.shift_counts = {}
        self.preassigned_shifts = {}
        self.exclusions = []
        self.schedule = {}

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

        self.update_dates()  # 在開啟時初始化

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

        self.btn_generate_schedule = ttk.Button(root, text="生成班表", command=self.generate_schedule)
        self.btn_generate_schedule.grid(row=10, column=1, padx=5, pady=5)

        self.btn_save_excel = ttk.Button(root, text="儲存 Excel", command=self.save_excel)
        self.btn_save_excel.grid(row=10, column=2, padx=5, pady=5)

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

    def update_dates(self, event=None):
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        _, num_days = calendar.monthrange(year, month)
        self.date_options = [str(day) for day in range(1, num_days + 1)]
        self.combo_date["values"] = self.date_options
        self.combo_date.set('')

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
        # 檢查該日期和班次是否已經有人排班
        if any(item['date'] == date and item['shift'] == shift_type for emp_shifts in self.preassigned_shifts.values() for item in emp_shifts):
            messagebox.showwarning("錯誤", f"{date} 的 {shift_type} 班次已經有人排班")
            return
        # 檢查員工是否在同一天排過其他班次
        if any(item['date'] == date for item in self.preassigned_shifts[employee]):
            messagebox.showwarning("錯誤", f"{employee} 已在 {date} 排過班")
            return
        # 不能連續兩天上班
        previous_day = str(int(date) - 1)
        next_day = str(int(date) + 1)
        if any(item['date'] == previous_day or item['date'] == next_day for item in self.preassigned_shifts[employee]):
            messagebox.showwarning("錯誤", f"{employee} 不能在 {date} 前後兩天排班")
            return
        # 員工班次超出限制
        if self.shift_counts[employee]["一線"] <= 0 and shift_type == "一線":
            messagebox.showwarning("錯誤", f"{employee} 的班次超出限制")
            return
        if self.shift_counts[employee]["二線"] <= 0 and shift_type == "二線":
            messagebox.showwarning("錯誤", f"{employee} 的班次超出限制")
            return
        # 將新的排班信息加入列表
        self.preassigned_shifts[employee].append({'date': date, 'shift': shift_type})
        messagebox.showinfo("成功", f"成功為 {employee} 預排 {date} 的 {shift_type} 班")
        self.update_preassigned_shifts_treeview()
        # 更新員工班次資訊
        if shift_type == "一線":
            self.shift_counts[employee]["一線"] -= 1
        elif shift_type == "二線":
            self.shift_counts[employee]["二線"] -= 1

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

    def generate_schedule(self):
        year = int(self.year_var.get())
        month = int(self.month_var.get())
        _, num_days = calendar.monthrange(year, month)
        # 初始化 schedule 和 last_assigned 記錄員工上次排班的日期
        self.schedule = {day: {"一線": None, "二線": None} for day in range(1, num_days + 1)}
        last_assigned = {e: -999 for e in self.employees.keys()}  # 初始值設為遠古時期
        remaining_shifts = {e: self.shift_counts[e]["一線"] + self.shift_counts[e]["二線"] for e in self.shift_counts}  # 剩餘需要排的班次數
        # 先將預先排班資料填入 schedule
        for employee, shifts in self.preassigned_shifts.items():
            for shift in shifts:
                day = int(shift['date'])
                shift_type = shift['shift']
                if day in self.schedule:
                    self.schedule[day][shift_type] = employee
        for day in range(1, num_days + 1):
            if self.schedule[day]["一線"] is not None and self.schedule[day]["二線"] is not None:
              continue  # 該日已排滿，跳過
            elif self.schedule[day]["一線"] is None and self.schedule[day]["二線"] is not None:
                available_primary = [e for e in self.employees.keys() if day not in self.unavailable_dates.get(e, [])]
                emp = self.schedule[day]["二線"]
                available_primary = [e for e in available_primary if e != emp and sorted([e, emp]) not in self.exclusions]  # 移除自己和 exclusions
            elif self.schedule[day]["一線"] is not None and self.schedule[day]["二線"] is None:
                available_secondary = [e for e in self.employees.keys() if day not in self.unavailable_dates.get(e, [])]
                emp = self.schedule[day]["一線"]
                available_secondary = [e for e in available_secondary if e != emp and sorted([e, emp]) not in self.exclusions]
            else:
                available_primary = [e for e in self.employees.keys() if day not in self.unavailable_dates.get(e, [])]
                available_secondary = available_primary.copy()

            # 移除前一天已經上班的員工
            if day > 1:
                prev_primary = self.schedule[day - 1]["一線"]
                prev_secondary = self.schedule[day - 1]["二線"]
                available_primary = [e for e in available_primary if e != prev_primary and e != prev_secondary]
                available_secondary = [e for e in available_secondary if e != prev_primary and e != prev_secondary]
            # 移除後一天已經上班的員工
            if day < num_days:
                after_primary = self.schedule[day + 1]["一線"]
                after_secondary = self.schedule[day + 1]["二線"]
                available_primary = [e for e in available_primary if e != after_primary and e != after_secondary]
                available_secondary = [e for e in available_secondary if e != after_primary and e != after_secondary]

            

            
            # 如果當天的班次已被預先指定，則跳過自動指派
            if self.schedule[day]["一線"] is not None:
                if self.schedule[day]["一線"] in available_primary:
                    available_primary.remove(self.schedule[day]["一線"])
            if self.schedule[day]["二線"] is not None:
                if self.schedule[day]["二線"] in available_secondary:
                    available_secondary.remove(self.schedule[day]["二線"])





            if not available_primary or not available_secondary:
                continue  # 如果當天沒人可排，跳過
            # 計算加權排序值：剩餘天數 / 尚須排班次數
            days_remaining = num_days - day + 1
            available_primary.sort(key=lambda e: (days_remaining / (remaining_shifts[e] + 1), last_assigned[e]))
            primary = available_primary[0] if self.schedule[day]["一線"] is None else self.schedule[day]["一線"]
            # 確保 secondary 不與 primary 組成禁排組合
            available_secondary = [e for e in available_secondary if e != primary and sorted([primary, e]) not in self.exclusions]
            available_secondary.sort(key=lambda e: (days_remaining / (remaining_shifts[e] + 1), last_assigned[e]))
            secondary = available_secondary[0] if available_secondary and self.schedule[day]["二線"] is None else self.schedule[day]["二線"]
            self.schedule[day]["一線"] = primary
            self.schedule[day]["二線"] = secondary
            last_assigned[primary] = day  # 更新排班記錄
            last_assigned[secondary] = day  # 更新排班記錄
            remaining_shifts[primary] -= 1
            remaining_shifts[secondary] -= 1
        messagebox.showinfo("成功", "班表已生成！")

    def save_excel(self):
        if not self.schedule:
            messagebox.showwarning("警告", "請先生成班表！")
            return
        # 將班表轉換為 DataFrame
        data = []
        for day, shifts in self.schedule.items():
            data.append([f"{self.year_var.get()}-{self.month_var.get()}-{day}", shifts["一線"], shifts["二線"]])
        df = pd.DataFrame(data, columns=["日期", "一線", "二線"])
        # 讓使用者選擇儲存位置
        filename = f"{self.year_var.get()}-{self.month_var.get()}"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=filename,
            filetypes=[("Excel 文件", "*.xlsx")],
            title="儲存 Excel"
        )
        if file_path:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"班表已儲存！")

if __name__ == "__main__":
    root = tk.Tk()
    app = SchedulingApp(root)
    root.mainloop()