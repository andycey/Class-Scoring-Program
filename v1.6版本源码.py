import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from datetime import datetime
import os
import pandas as pd
import tkinter.font as font
import decimal
from datetime import datetime, timedelta
from tkinter import ttk
class ClassScoringApp:
    def __init__(self, root):
        self.root = root
        self.root.title("v1.6班级计分程序")
        self.root.geometry("500x700")  # 增加窗口尺寸以适应新布局

        # 设置样式
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f0f0f0')
        self.style.configure('Header.TLabel', background='#8fd3f4', foreground='white', font=('宋体', 14, 'bold'))
        self.style.configure('TButton', font=('Arial', 10), padding=5)
        self.style.configure('TLabel', background='#f0f0f0', font=('Arial', 10))
        self.style.configure('TEntry', font=('Arial', 10))
        self.style.configure('TOptionMenu', font=('Arial', 10))

        self.students = self.load_students()
        self.groups = self.load_groups()
        self.scores = self.load_scores()
        self.undo_stack = []  # 撤销栈
        self.redo_stack = []  # 重做栈

        self.create_widgets()

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = 500
        window_height = 700
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # 设置背景色
        root.configure(background='#f0f0f0')

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 标题
        header = ttk.Label(main_frame, text="班级计分系统", style='Header.TLabel',font=('宋体', 20))
        header.grid(row=0, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))

        # 输入区域
        input_frame = ttk.LabelFrame(main_frame, text="计分输入", padding="10")
        input_frame.grid(row=1, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        input_frame.columnconfigure(1, weight=1)

        # 姓名输入
        ttk.Label(input_frame, text="姓名:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(input_frame)
        self.name_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 类型选择
        ttk.Label(input_frame, text="类型:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.type_var = tk.StringVar()
        type_options = ["作业加分", "作业扣分", "纪律", "班委加分", "早读优秀", "班务", "黑板报", "卫生", "其他"]
        self.type_menu = ttk.OptionMenu(input_frame, self.type_var, "纪律", *type_options)
        self.type_menu.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 原因输入
        ttk.Label(input_frame, text="原因:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.reason_entry = ttk.Entry(input_frame)
        self.reason_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 分数输入
        ttk.Label(input_frame, text="分数:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.score_entry = ttk.Entry(input_frame)
        self.score_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=5)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)

        # 第一行按钮
        self.add_button = ttk.Button(button_frame, text="添加分数", command=self.add_score)
        self.add_button.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.view_button = ttk.Button(button_frame, text="查看分数", command=self.view_scores)
        self.view_button.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

        # 第二行按钮
        self.undo_button = ttk.Button(button_frame, text="撤销", command=self.undo_action)
        self.undo_button.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))

        self.redo_button = ttk.Button(button_frame, text="重做", command=self.redo_action)
        self.redo_button.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))

        # 功能区域
        function_frame = ttk.LabelFrame(main_frame, text="功能选项", padding="10")
        function_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky=(tk.W, tk.E))
        function_frame.columnconfigure(0, weight=1)

        # 功能按钮
        functions = [
            ("导入作业统计表", self.import_homework),
            ("个人排序", self.sort_students_window),
            ("查看小组平均分", self.view_group_average_window),
            ("按小组批量加减分", self.batch_scoring_window),
            ("导出数据到Excel", self.export_to_excel),
            ("生成作业一周报表", self.generate_homework_report),
            ("更新日志", self.show_update_log)
        ]

        for i, (text, command) in enumerate(functions):
            btn = ttk.Button(function_frame, text=text, command=command)
            btn.grid(row=i, column=0, pady=5, sticky=(tk.W, tk.E))

        # 添加按 Enter 键计分的事件绑定
        self.root.bind('<Return>', lambda event: self.add_score())
    def load_students(self):
        try:
            with open("班级名单.txt", "r", encoding="utf-8") as file:
                students = file.read().splitlines()
            return students
        except FileNotFoundError:
            messagebox.showerror("错误", "班级名单.txt 文件未找到")
            self.root.destroy()

    def import_homework(self):
        try:
            file_path = filedialog.askopenfilename(title="选择作业统计表", filetypes=[("Excel 文件", "*.xlsx")])
            if not file_path:
                return

            subject = os.path.splitext(os.path.basename(file_path))[0]  # 提取科目名称
            df = pd.read_excel(file_path, header=0)
            current_year = datetime.now().year

            # 第一列为编号，第二列为学生姓名，从第三列开始为日期
            for col in df.columns[2:]:  # 从第三列开始
                date = col
                if isinstance(date, (str, float)):  # 确保列值类型可解析
                    date = str(date)  # 转换为字符串
                    if "." in date:
                        month, day = map(int, date.split("."))
                        full_date = datetime(current_year, month, day).strftime("%Y-%m-%d")
                    else:
                        print(f"跳过列 {col} - 无法解析为日期")
                        continue
                else:
                    print(f"跳过列 {col} - 非日期列")
                    continue

                # 遍历学生和对应日期列状态
                for student, status in zip(df.iloc[:, 1], df[col]):
                    if student not in self.students:
                        print(f"跳过学生 {student} - 不在班级名单中")
                        continue

                    if str(status).strip() == "X":  # 确保状态是字符串
                        reason = f"{subject}作业未完成"
                        score = -0.5  # 默认扣分值

                        if student not in self.scores:
                            self.scores[student] = []

                        entry = f"{full_date}  作业 - {reason}  {score}"
                        self.scores[student].append(entry)
                        print(f"{student}记录已添加: {entry}")

            self.update_scores_file()
            messagebox.showinfo("反馈", "作业统计表已导入并处理完毕")
        except Exception as e:
            messagebox.showerror("错误", f"导入作业统计表时出错: {str(e)}")
    def load_groups(self):
        try:
            with open("分组.txt", "r", encoding="utf-8") as file:
                groups = file.read().splitlines()
            return groups
        except FileNotFoundError:
            messagebox.showerror("错误", "分组.txt 文件未找到")
            self.root.destroy()

    def load_scores(self):
        scores = {}
        if os.path.exists("积分.txt"):
            with open("积分.txt", "r", encoding="utf-8") as file:
                lines = file.read().splitlines()

            current_name = None
            for line in lines:
                if line.endswith(":"):
                    current_name = line[:-1]
                    scores[current_name] = []
                elif current_name:
                    scores[current_name].append(line)

        return scores

    def add_score(self):
        name = self.name_entry.get()
        score_type = self.type_var.get()
        reason = self.reason_entry.get()
        score_str = self.score_entry.get()

        if not name or not score_type or not reason or not score_str:
            messagebox.showerror("错误", "请填写完整信息")
            return

        try:
            score = float(score_str)
            score = decimal.Decimal(score)
        except ValueError:
            messagebox.showerror("错误", "分数必须是数字")
            return

        if name not in self.students:
            messagebox.showerror("错误", f"{name} 不在班级名单中")
            return

        if name not in self.scores:
            self.scores[name] = []

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"{timestamp}  {score_type} - {reason}  {score}"

        # 保存操作到撤销栈
        self.undo_stack.append((name, entry))
        self.redo_stack.clear()  # 新操作会清空重做栈

        self.scores[name].append(entry)
        self.update_scores_file()

        # 自动清空姓名输入框
        self.name_entry.delete(0, tk.END)

        messagebox.showinfo("反馈", f"{name} 的分数已添加")

    def undo_action(self):
        if not self.undo_stack:
            messagebox.showinfo("提示", "没有可撤销的操作")
            return

        name, entry = self.undo_stack.pop()
        if name in self.scores and entry in self.scores[name]:
            self.scores[name].remove(entry)
            self.redo_stack.append((name, entry))  # 保存操作到重做栈
            self.update_scores_file()
            messagebox.showinfo("撤销操作", f"撤销了：\n{name}{entry}")

    def redo_action(self):
        if not self.redo_stack:
            messagebox.showinfo("提示", "没有可重做的操作")
            return

        name, entry = self.redo_stack.pop()
        self.scores[name].append(entry)
        self.undo_stack.append((name, entry))  # 保存操作到撤销栈
        self.update_scores_file()
        messagebox.showinfo("重做操作", f"重做了：\n{entry}")
    def update_scores_file(self):
        with open("积分.txt", "w", encoding="utf-8") as file:
            for name, entries in self.scores.items():
                file.write(f"{name}:\n")
                for entry in entries:
                    file.write(f"{entry}\n")

    def view_group_average_window(self):
        # 将所有小组按照平均分降序排列
        sorted_groups = sorted(self.groups, key=lambda x: self.calculate_group_average_score(x), reverse=True)

        average_window = tk.Toplevel(self.root)
        average_window.title("所有小组分数详情")

        canvas = tk.Canvas(average_window)
        canvas.pack(side=tk.LEFT, fill=tk.Y)

        scrollbar = tk.Scrollbar(average_window, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor=tk.NW)

        for group_line in sorted_groups:
            group_info = group_line.split("：")
            group_name = group_info[0]
            group_members = group_info[1].split(" ")
            group_average = self.calculate_group_average_score(group_line)

            label = tk.Label(frame, text=f"{group_name} 平均分: {group_average:.2f}")
            label.pack()

            members_str = ""
            for member in group_members:
                if member in self.scores:
                    member_total_score = sum(float(entry.split()[-1]) for entry in self.scores[member])
                    members_str += f"{member}: {member_total_score}\n"

            members_label = tk.Label(frame, text=members_str)
            members_label.pack()

        frame.update_idletasks()

        canvas.config(scrollregion=canvas.bbox("all"))

    def calculate_group_average_score(self, group_line):
        group_info = group_line.split("：")
        group_members = group_info[1].split(" ")
        total_score = 0

        for member in group_members:
            if member in self.scores:
                member_scores = [float(entry.split()[-1]) for entry in self.scores[member]]
                member_total_score = sum(member_scores)
                total_score += member_total_score

        group_average = total_score / len(group_members) if len(group_members) > 0 else 0
        return group_average

    def view_scores(self):
        name = self.name_entry.get()
        if name not in self.scores or not self.scores[name]:
            messagebox.showinfo("提示", f"{name}暂无积分记录")
            return

        scores_str = "\n".join(self.scores[name])
        messagebox.showinfo(f"{name}的积分记录", scores_str)
    def sort_students_window(self):
        sorted_students = sorted(self.students, key=lambda x: self.calculate_total_score(x), reverse=True)

        sort_window = tk.Toplevel(self.root)
        sort_window.title("个人排序")
        sort_window.geometry("900x600")  # 设置排序窗口尺寸

        # 设置窗口出现在屏幕正中央
        screen_width = sort_window.winfo_screenwidth()
        screen_height = sort_window.winfo_screenheight()
        window_width = 800
        window_height = 600
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        sort_window.geometry(f"{window_width}x{window_height}+{x}+{y}")

        canvas = tk.Canvas(sort_window)
        canvas.pack(side=tk.LEFT, fill=tk.Y)

        scrollbar = tk.Scrollbar(sort_window, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor=tk.NW)

        for i, student in enumerate(sorted_students, start=1):
            total_score = self.calculate_total_score(student)
            label_text = f"{i}. {student}: {total_score}"

            # 如果积分未变动，则添加标记
            if student not in self.scores:
                label_text += " (未变动)"

            label = tk.Label(frame, text=label_text, font=("宋体", 30))  # 修改字号为30号
            label.pack()

        frame.update_idletasks()

        canvas.config(scrollregion=canvas.bbox("all"))

    def show_update_log(self):
        messagebox.showinfo("1.6更新日志","1.修复了导出到excel时个人排序不显示没有积分加扣分记录的同学的bug。\n2.优化了UI界面，使UI更加简洁明了。")


    def batch_scoring_window(self):
        batch_window = tk.Toplevel(self.root)
        batch_window.title("按分组批量加减分")


        self.batch_var = tk.IntVar(value=0)  # 默认选择加分
        batch_radio_add = tk.Radiobutton(batch_window, text="加分", variable=self.batch_var, value=0)
        batch_radio_add.pack()
        batch_radio_subtract = tk.Radiobutton(batch_window, text="减分", variable=self.batch_var, value=1)
        batch_radio_subtract.pack()

        self.group_var = tk.StringVar(value=self.groups[0])  # 默认选择第一个小组
        batch_group_menu = tk.OptionMenu(batch_window, self.group_var, *self.groups)
        batch_group_menu.pack()

        batch_button = tk.Button(batch_window, text="批量操作", command=self.perform_batch_scoring)
        batch_button.pack()

    def generate_homework_report(self):
        try:
            # 用户输入当周应交作业数
            num_assignments = int(simpledialog.askstring("输入", "请输入当周应交作业数："))
            if num_assignments <= 0:
                messagebox.showerror("错误", "应交作业数必须大于零")
                return

            # 获取最近一周的日期范围
            today = datetime.today()
            last_week = today - timedelta(days=7)
            last_week_str = last_week.strftime("%Y-%m-%d")

            report_data = []
            detail_data = []  # 用来存储每个学生的作业未完成详情

            # 遍历学生成绩记录
            for name, entries in self.scores.items():
                missed_count = 0
                total_count = 0
                details = []  # 确保 details 是列表类型，用于存储未交作业的详情

                for entry in entries:
                    try:
                        parts = entry.split("  ")
                        timestamp, details_str = parts[0], parts[1]
                        score_type, reason_score = details_str.split(" - ")
                        timestamp = timestamp.split(" ")[0]
                        # 检查记录是否为作业扣分类型，并在最近一周内
                        if timestamp >= last_week_str and "作业扣分" in score_type:
                            total_count += 1
                            missed_count += 1
                                # 添加未完成作业的详情
                            details.append(f"{timestamp} - {reason_score}")
                            print(details)
                    except ValueError:
                        print(f"跳过无效记录: {entry}")
                        continue

                # 添加到报表汇总数据
                completion_rate = (
                            (num_assignments - missed_count) / num_assignments * 100) if num_assignments > 0 else 100
                report_data.append([name, num_assignments, missed_count, completion_rate])

                # 添加未交作业详情到明细数据
                detail_data.append([name, "; ".join(details) if details else "无未完成记录"])

            # 创建 DataFrame
            summary_df = pd.DataFrame(report_data, columns=["姓名", "应交作业数", "未完成作业数", "完成率"])
            details_df = pd.DataFrame(detail_data, columns=["姓名", "未交作业详情"])

            # 导出到 Excel
            with pd.ExcelWriter("作业一周报表.xlsx") as writer:
                summary_df.to_excel(writer, index=False, sheet_name="汇总统计")
                details_df.to_excel(writer, index=False, sheet_name="未完成详情")

            messagebox.showinfo("反馈", "作业一周报表已生成并导出")
        except Exception as e:
            messagebox.showerror("错误", f"生成作业一周报表时出错: {str(e)}")



    def calculate_total_score(self, student):
        if student in self.scores:
            total_score = sum(float(entry.split()[-1]) for entry in self.scores[student])
            return total_score
        else:
            return 0

    def calculate_total_score(self, student):
        if student in self.scores:
            total_score = sum(float(entry.split()[-1]) for entry in self.scores[student])
            return total_score
        else:
            return 0

    # 原batch_scoring_window需要替换为以下内容（添加窗口居中设置）
    def batch_scoring_window(self):
        batch_group_window = tk.Toplevel(self.root)
        batch_group_window.title("选择小组")

        # 设置窗口居中
        batch_group_window.update_idletasks()
        width = 300
        height = 200
        x = (batch_group_window.winfo_screenwidth() // 2) - (width // 2)
        y = (batch_group_window.winfo_screenheight() // 2) - (height // 2)
        batch_group_window.geometry(f'{width}x{height}+{x}+{y}')

        # 按钮让用户选择小组文件
        def select_group_file():
            file_path = filedialog.askopenfilename(title="选择小组文件", filetypes=[("文本文件", "*.txt")])
            if file_path:
                self.group_file_path = file_path
                messagebox.showinfo("选择文件", f"已选择小组文件: {file_path}")
                self.show_group_selection_window(file_path)

        select_button = tk.Button(batch_group_window, text="选择小组文件", command=select_group_file)
        select_button.pack(pady=20)

    def show_group_selection_window(self, file_path):
        try:
            # 读取小组文件
            with open(file_path, "r", encoding="utf-8") as file:
                groups = file.readlines()

            # 提取小组信息
            self.groups = {}
            for group_line in groups:
                group_info = group_line.strip().split("：")
                group_name = group_info[0]
                members = group_info[1].split(" ")
                self.groups[group_name] = members  # 将小组名和成员存储到字典中

            # 创建一个新窗口来选择小组
            group_selection_window = tk.Toplevel(self.root)
            group_selection_window.title("选择小组")

            # 创建小组选择菜单
            self.group_var = tk.StringVar(value=list(self.groups.keys())[0])  # 默认选择第一个小组
            batch_group_menu = tk.OptionMenu(group_selection_window, self.group_var, *self.groups.keys(),
                                             command=self.update_members_label)
            batch_group_menu.pack()

            # 选择并执行批量加扣分操作的按钮
            select_button = tk.Button(group_selection_window, text="选择并执行", command=self.show_batch_entry_window)
            select_button.pack()

            # 显示扣分类型选择
            self.deduction_type_var = tk.StringVar(value="作业扣分")  # 默认选择作业扣分
            batch_deduction_type_menu = tk.OptionMenu(group_selection_window, self.deduction_type_var, "作业加分","作业扣分", "纪律","班委加分", "早读优秀", "班务", "黑板报",
                                       "卫生", "其他")
            batch_deduction_type_menu.pack()

            # 显示选中的小组成员
            self.members_label = tk.Label(group_selection_window,
                                          text="成员列表: " + ", ".join(self.groups[self.group_var.get()]))
            self.members_label.pack()

            # 按 Enter 键执行批量操作
            group_selection_window.bind('<Return>', lambda event: self.show_batch_entry_window())

        except Exception as e:
            messagebox.showerror("错误", f"读取小组文件时出错: {str(e)}")

    def update_members_label(self, selected_group):
        # 更新显示的成员列表
        members = self.groups[selected_group]
        self.members_label.config(text="成员列表: " + ", ".join(members))

    def show_batch_entry_window(self):
        if not hasattr(self, 'group_file_path'):
            messagebox.showerror("错误", "请先选择小组文件")
            return

        # 获取一次用户输入的分数和原因
        score_str = simpledialog.askstring("输入分数", f"请输入批量操作的分数:")
        if not score_str:
            return

        reason = simpledialog.askstring("输入原因", f"请输入批量操作的原因:")
        if not reason:
            return

        # 获取扣分类型
        deduction_type = self.deduction_type_var.get()

        try:
            # 获取选择的小组成员
            selected_group = self.group_var.get()
            members = self.groups[selected_group]

            scores_to_add = {}

            # 为每个成员生成记录
            for member in members:
                name = member.strip()
                if name not in scores_to_add:
                    scores_to_add[name] = []

                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # 添加时分秒
                entry = f"{timestamp}  {deduction_type} - {reason}  {score_str}"
                scores_to_add[name].append(entry)

            # 将分数添加到每个学生的记录中
            for name, entries in scores_to_add.items():
                if name not in self.scores:
                    self.scores[name] = []

                self.scores[name].extend(entries)

            self.update_scores_file()
            messagebox.showinfo("反馈", f"批量操作完成")
        except Exception as e:
            messagebox.showerror("错误", f"读取小组文件时出错: {str(e)}")

    def export_to_excel(self):
        df = pd.DataFrame(columns=["姓名", "日期", "类型", "原因", "分数"])
        ef = pd.DataFrame()

        # 创建包含所有学生的DataFrame，初始分数为0
        all_students_df = pd.DataFrame({"姓名": self.students})
        all_students_df["分数"] = 0.0

        for name, entries in self.scores.items():
            for entry in entries:
                parts = entry.split("  ")
                timestamp, details = parts[0], parts[1]
                score = parts[2]  # 修正：加扣分值应在 parts[2]
                date = timestamp.split(" ")[0]  # 提取日期部分
                date += ' '
                score_type, reason = details.split(" - ")
                df = pd.concat([df, pd.DataFrame({
                    "姓名": [name],
                    "日期": [date],
                    "类型": [score_type],
                    "原因": [reason],
                    "分数": [float(score)]
                })])

        # 计算每个学生的总分
        if not df.empty:
            student_scores = df.groupby("姓名")["分数"].sum().reset_index()
            # 更新所有学生DataFrame中的分数
            all_students_df = all_students_df.merge(student_scores, on="姓名", how="left", suffixes=('_old', ''))
            all_students_df["分数"] = all_students_df["分数"].fillna(all_students_df["分数_old"])
            all_students_df = all_students_df[["姓名", "分数"]]

        # 使用rank方法计算排名，相同分数排名相同
        all_students_df['排行'] = all_students_df['分数'].rank(method='min', ascending=False).astype(int)
        all_students_df = all_students_df.sort_values(by=['排行', '姓名'])

        # 在 Sheet1 中导出详细记录
        with pd.ExcelWriter("班级积分数据.xlsx") as writer:
            df.to_excel(writer, index=False, sheet_name="详细记录")

            # 在 Sheet2 中导出各类型总分并排序
            if not df.empty:
                summary = df.groupby(["姓名", "类型"]).sum().reset_index()
                summary = summary.sort_values(by=["类型", "分数"], ascending=[True, True])
                summary.to_excel(writer, index=False, sheet_name="类型统计")
            else:
                pd.DataFrame(columns=["姓名", "类型", "分数"]).to_excel(writer, index=False, sheet_name="类型统计")

            # 在 Sheet3 中导出总分并排序（包含所有学生）
            total_scores = all_students_df[['排行', '姓名', '分数']]
            total_scores.to_excel(writer, index=False, sheet_name="总分排序")

            # 在sheet4中导出小组平均分并排列
            group_avg_data = []
            for group_line in self.groups:
                group_info = group_line.split("：")
                group_name = group_info[0]
                group_members = group_info[1].split(" ")

                total_score = 0
                member_count = 0

                for member in group_members:
                    member = member.strip()
                    if member in all_students_df['姓名'].values:
                        member_score = all_students_df[all_students_df['姓名'] == member]['分数'].values[0]
                        total_score += member_score
                        member_count += 1

                        member_df = pd.DataFrame({
                            "小组": [group_name],
                            "小组成员": [member],
                            "个人分数": [member_score]
                        })
                        ef = pd.concat([ef, member_df])

                group_average = total_score / member_count if member_count > 0 else 0
                group_df = pd.DataFrame({
                    "小组": [group_name],
                    "小组平均分": [group_average],
                    "小组总人数": [member_count],
                    "小组总分": [total_score]
                })
                group_avg_data.append(group_df)

            # 添加小组平均分数据
            if group_avg_data:
                group_avg_df = pd.concat(group_avg_data)
                group_avg_df = group_avg_df.sort_values(by="小组平均分", ascending=False)
                group_avg_df.to_excel(writer, index=False, sheet_name="小组排序")

            # 添加小组成员详情
            if not ef.empty:
                ef.to_excel(writer, index=False, sheet_name="小组成员详情")

        messagebox.showinfo("反馈", "数据已导出到班级积分数据.xlsx")

if __name__ == "__main__":
    root = tk.Tk()
    app = ClassScoringApp(root)
    root.mainloop()
