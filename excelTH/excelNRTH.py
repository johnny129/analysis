import os
import win32com.client as win32
from ttkbootstrap import *
from tkinter import filedialog, messagebox
from datetime import datetime

class ExcelFindReplaceTool:
    def __init__(self, master):
        self.master = master
        self.master.title("Excel查找替换工具")
        self.master.geometry("800x600")

        self.excel = None
        self.workbook = None
        self.log = []

        self.create_widgets()
        self.connect_to_excel()

    def create_widgets(self):
        main_frame = Frame(self.master)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        input_frame = LabelFrame(main_frame, text="查找替换设置", bootstyle="info")
        input_frame.pack(fill=X, pady=5)

        # 查找内容
        Label(input_frame, text="查找内容:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.find_entry = Entry(input_frame, width=40)
        self.find_entry.grid(row=0, column=1, padx=5, pady=5, columnspan=2)

        # 替换内容
        Label(input_frame, text="替换内容:").grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.replace_entry = Entry(input_frame, width=40)
        self.replace_entry.grid(row=1, column=1, padx=5, pady=5, columnspan=2)

        # 转义字符开关
        self.escape_var = BooleanVar(value=True)
        Checkbutton(input_frame,
                   text="启用转义字符（\\n 表示换行）",
                   variable=self.escape_var,
                   bootstyle="info-roundtoggle").grid(row=2, column=0, columnspan=3, sticky=W, padx=5)

        # 查找范围
        Label(input_frame, text="查找范围:").grid(row=3, column=0, padx=5, pady=5, sticky=W)
        self.scope_var = StringVar(value="worksheet")
        Radiobutton(input_frame, text="当前工作表", variable=self.scope_var, value="worksheet").grid(row=3, column=1, padx=5, sticky=W)
        Radiobutton(input_frame, text="整个工作簿", variable=self.scope_var, value="workbook").grid(row=3, column=2, padx=5, sticky=W)

        # 查找目标
        Label(input_frame, text="查找目标:").grid(row=4, column=0, padx=5, pady=5, sticky=W)
        self.cell_var = BooleanVar(value=True)
        self.textbox_var = BooleanVar(value=True)
        Checkbutton(input_frame, text="单元格", variable=self.cell_var).grid(row=4, column=1, sticky=W)
        Checkbutton(input_frame, text="文本框/图形", variable=self.textbox_var).grid(row=4, column=2, sticky=W)

        # 功能按钮
        btn_frame = Frame(main_frame)
        btn_frame.pack(pady=10)
        Button(btn_frame, text="开始替换", command=self.start_replace, bootstyle="success").pack(side=LEFT, padx=5)
        Button(btn_frame, text="重新连接Excel", command=self.reconnect_excel, bootstyle="secondary").pack(side=LEFT, padx=5)
        Button(btn_frame, text="清除日志", command=self.clear_log, bootstyle="warning").pack(side=LEFT, padx=5)

        # 日志区域
        log_frame = LabelFrame(main_frame, text="操作日志", bootstyle="primary")
        log_frame.pack(fill=BOTH, expand=True)

        self.log_text = Text(log_frame, wrap=WORD, height=10)
        scrollbar = Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.pack(fill=BOTH, expand=True)

    def reconnect_excel(self):
        self.connect_to_excel(force=True)

    def connect_to_excel(self, force=False):
        try:
            if force or not self.excel:
                try:
                    self.excel = win32.GetActiveObject("Excel.Application")
                    self.log_action("成功连接到正在运行的Excel实例")
                except:
                    self.excel = win32.gencache.EnsureDispatch("Excel.Application")
                    self.excel.Visible = True
                    self.log_action("已创建新的Excel实例")

            if self.excel.Workbooks.Count > 0:
                try:
                    self.workbook = self.excel.ActiveWorkbook
                    self.log_action(f"已连接工作簿: {self.workbook.Name}")
                except:
                    self.workbook = self.excel.Workbooks[0]
                    self.log_action(f"已连接第一个打开的工作簿: {self.workbook.Name}")
            else:
                self.workbook = None
                self.log_action("检测到Excel程序，但未打开任何工作簿")
                self.prompt_open_file()
        except Exception as e:
            self.show_connection_error(f"连接失败: {str(e)}")

    def prompt_open_file(self):
        if messagebox.askyesno("工作簿未找到", "是否要手动选择Excel文件？"):
            file_path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xls *.xlsx *.xlsm")])
            if file_path:
                try:
                    self.workbook = self.excel.Workbooks.Open(os.path.abspath(file_path))
                    self.excel.Visible = True
                    self.log_action(f"已打开工作簿: {self.workbook.Name}")
                except Exception as e:
                    messagebox.showerror("错误", f"文件打开失败: {str(e)}")
            else:
                self.log_action("用户取消文件选择")

    def validate_connection(self):
        try:
            if not self.excel or not self.workbook:
                return False
            self.excel.Name
            self.workbook.Name
            return True
        except:
            self.log_action("连接已失效，正在尝试重新连接...")
            self.connect_to_excel(force=True)
            return self.workbook is not None

    def process_escape_chars(self, text):
        if self.escape_var.get():
            return text.replace(r'\n', '\n').replace(r'\r', '\r')
        return text

    def process_worksheet(self, sheet):
        replacements = 0

        find_text = self.process_escape_chars(self.find_entry.get())
        replace_text = self.process_escape_chars(self.replace_entry.get())

        # 处理单元格
        if self.cell_var.get():
            try:
                used_range = sheet.UsedRange
                values = used_range.Value
                if values:
                    new_values = []
                    modified = False
                    for row in values:
                        new_row = []
                        for cell_value in row:
                            if isinstance(cell_value, str) and find_text in cell_value:
                                new_value = cell_value.replace(find_text, replace_text)
                                new_row.append(new_value)
                                replacements += cell_value.count(find_text)
                                modified = True
                            else:
                                new_row.append(cell_value)
                        new_values.append(tuple(new_row))
                    if modified:
                        used_range.Value = tuple(new_values)
                        self.log_action(f"工作表 [{sheet.Name}]：单元格共替换 {replacements} 项")
            except Exception as e:
                self.log_action(f"处理单元格出错：{e}")

        # 处理图形和文本框
        if self.textbox_var.get():
            def process_shape(shape, sheet_name):
                nonlocal replacements
                if shape.Type == 17:
                    tf = shape.TextFrame2
                elif shape.Type == 1:
                    tf = shape.TextFrame2
                    if not tf.HasText:
                        return
                elif shape.Type == 6:
                    for sub in shape.GroupItems:
                        process_shape(sub, sheet_name)
                    return
                else:
                    return

                text_range = tf.TextRange
                orig = text_range.Text
                if find_text in orig:
                    newt = orig.replace(find_text, replace_text)
                    text_range.Text = newt
                    self.log_action(f"工作表 [{sheet_name}] 形状 [{shape.Name}]：{orig} → {newt}")
                    replacements += orig.count(find_text)

            for shp in sheet.Shapes:
                process_shape(shp, sheet.Name)

            # 处理ActiveX控件
            try:
                for ole_obj in sheet.OLEObjects:
                    obj = ole_obj.Object
                    if hasattr(obj, 'Text'):
                        orig_text = obj.Text
                        if find_text in orig_text:
                            new_text = orig_text.replace(find_text, replace_text)
                            obj.Text = new_text
                            self.log_action(f"工作表 [{sheet.Name}] ActiveX控件 [{ole_obj.Name}]：{orig_text} → {new_text}")
                            replacements += orig_text.count(find_text)
                    elif hasattr(obj, 'Caption'):
                        orig_caption = obj.Caption
                        if find_text in orig_caption:
                            new_caption = orig_caption.replace(find_text, replace_text)
                            obj.Caption = new_caption
                            self.log_action(f"工作表 [{sheet.Name}] ActiveX标签 [{ole_obj.Name}]：{orig_caption} → {new_caption}")
                            replacements += orig_caption.count(find_text)
            except Exception as e:
                self.log_action(f"处理ActiveX控件出错：{str(e)}")

        return replacements

    def start_replace(self):
        if not self.validate_connection():
            messagebox.showwarning("警告", "未连接到有效的Excel工作簿")
            return
        if not self.find_entry.get().strip():
            messagebox.showwarning("警告", "查找内容不能为空")
            return

        try:
            self.excel.ScreenUpdating = False
            self.excel.Calculation = -4135  # xlCalculationManual

            total = 0
            sheets = [self.excel.ActiveSheet] if self.scope_var.get() == "worksheet" else list(self.workbook.Sheets)
            for sht in sheets:
                total += self.process_worksheet(sht)

            self.workbook.Save()
            self.log_action(f"操作完成！共完成 {total} 处替换")
            self.generate_log_file()
        except Exception as e:
            messagebox.showerror("错误", f"操作失败: {e}")
            self.log_action(f"错误发生: {e}")
        finally:
            self.excel.ScreenUpdating = True
            self.excel.Calculation = -4105  # xlCalculationAutomatic
            self.update_log()

    def log_action(self, msg):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log.append(f"[{ts}] {msg}")
        self.update_log()

    def update_log(self):
        self.log_text.delete(1.0, END)
        self.log_text.insert(END, "\n".join(self.log[-50:]))
        self.log_text.see(END)

    def clear_log(self):
        self.log = []
        self.update_log()

    def generate_log_file(self):
        log_dir = os.path.join(os.getcwd(), "operation_logs")
        os.makedirs(log_dir, exist_ok=True)
        fname = datetime.now().strftime("replace_log_%Y%m%d_%H%M%S.txt")
        path = os.path.join(log_dir, fname)
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log))
        self.log_action(f"日志文件已保存至: {path}")

    def show_connection_error(self, message):
        err = f"""{message}

可能原因：
1. 未安装Microsoft Excel
2. Excel文件未打开或已被关闭
3. 文件被其他进程锁定
4. 权限不足
5. 存在多个Excel实例

解决方法：
1. 确保Excel已安装并打开文件
2. 以管理员身份运行本程序
3. 关闭可能锁定文件的程序
4. 检查Excel的COM设置
5. 关闭多余的Excel实例"""
        messagebox.showerror("连接错误", err)
        self.master.destroy()

if __name__ == "__main__":
    app = Window(title="Excel查找替换工具", themename="litera")
    ExcelFindReplaceTool(app)
    app.mainloop()