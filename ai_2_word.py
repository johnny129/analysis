import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import comtypes.client

# --------------------------
# AI 文件处理核心逻辑
# --------------------------
def extract_text_from_ai(ai_file, merge_segments=False, merge_threshold=50):
    """
    从 AI 文件中提取所有文本
    :param ai_file: AI 文件路径
    :param merge_segments: 是否合并相邻句段
    :param merge_threshold: 合并的最大垂直距离阈值
    :return: 文本列表
    """
    try:
        ai = comtypes.client.CreateObject("Illustrator.Application")
        doc = ai.Open(ai_file)
        text_frames = []
        
        # 收集所有文本框及其位置
        for text_frame in doc.TextFrames:
            content = text_frame.Contents
            position = text_frame.Position
            # 获取文本框的边界框
            try:
                # 尝试获取几何边界
                geometric_bounds = text_frame.GeometricBounds
                height = abs(geometric_bounds[0] - geometric_bounds[2])
            except:
                # 如果无法获取几何边界，使用默认高度
                height = 20
                
            text_frames.append({
                'content': content,
                'x': position[0],
                'y': position[1],
                'height': height
            })
        
        doc.Close()
        
        # 如果需要合并句段
        if merge_segments:
            return merge_adjacent_segments(text_frames, merge_threshold)
        else:
            return [frame['content'] for frame in text_frames]
        
    except Exception as e:
        messagebox.showerror("错误", f"无法提取文本: {str(e)}")
        return []

def merge_adjacent_segments(text_frames, threshold=50):
    """
    合并相邻的文本段
    :param text_frames: 文本帧列表
    :param threshold: 合并的最大垂直距离阈值
    :return: 合并后的文本列表
    """
    # 按垂直位置排序（从上到下）
    sorted_frames = sorted(text_frames, key=lambda f: f['y'], reverse=True)
    
    merged_segments = []
    current_segment = None
    current_y = None
    
    for frame in sorted_frames:
        content = frame['content'].strip()
        y = frame['y']
        height = frame['height']
        
        # 如果是第一个句段
        if current_segment is None:
            current_segment = content
            current_y = y
            continue
        
        # 计算垂直距离
        vertical_distance = abs(current_y - y)
        
        # 检查是否在同一行或相邻行（考虑高度）
        if vertical_distance < (height * 1.5 + threshold):
            # 合并内容（添加空格）
            current_segment += " " + content
            # 更新当前y位置为合并后的平均位置
            current_y = (current_y + y) / 2
        else:
            # 保存当前合并的句段
            merged_segments.append(current_segment)
            # 开始新的句段
            current_segment = content
            current_y = y
    
    # 添加最后一个句段
    if current_segment:
        merged_segments.append(current_segment)
    
    return merged_segments

def generate_translation_csv(texts, output_csv, filename, export_numbers=True, export_blanks=True):
    """生成翻译用 CSV 文件（支持过滤纯数字和空白内容）"""
    try:
        with open(output_csv, 'w', encoding='utf-8-sig', newline='') as csvfile:
            writer = csv.writer(csvfile)
            # 写入文件名作为第一行
            writer.writerow([f"文件名: {filename}"])
            writer.writerow(["原文", "译文"])  # 表头
            
            # 应用过滤规则
            filtered_texts = []
            for text in texts:
                # 检查是否为空白内容
                is_blank = not text.strip()
                
                # 检查是否为纯数字
                is_numeric = False
                try:
                    # 尝试转换为数字（处理各种格式）
                    cleaned_text = text.replace(',', '').replace('.', '').replace(' ', '')
                    if cleaned_text.isdigit():
                        is_numeric = True
                except:
                    pass
                
                # 应用过滤规则
                if (is_blank and not export_blanks) or (is_numeric and not export_numbers):
                    continue  # 跳过不符合条件的文本
                
                filtered_texts.append(text)
            
            # 写入过滤后的文本
            for text in filtered_texts:
                writer.writerow([text, text])  # 原文和译文初始相同
        
        messagebox.showinfo("成功", f"翻译模板已生成: {output_csv}")
        return True
    except Exception as e:
        messagebox.showerror("错误", f"生成CSV失败: {str(e)}")
        return False

def update_ai_file(ai_file, translations, mode, font=None):
    """更新 AI 文件（替换或追加译文），并支持自定义字体设置"""
    try:
        ai = comtypes.client.CreateObject("Illustrator.Application")
        doc = ai.Open(ai_file)
        
        # 获取所有文本框
        text_frames = [frame for frame in doc.TextFrames]
        
        # 检查译文数量是否匹配
        if mode == "replace" and len(translations) != len(text_frames):
            messagebox.showwarning("警告", 
                f"译文数量({len(translations)})与文本框数量({len(text_frames)})不匹配！\n"
                "请确保导出和导入时使用了相同的合并设置。")
        
        for idx, text_frame in enumerate(text_frames):
            if idx >= len(translations):
                break  # 防止索引越界
                
            if mode == "replace":
                text_frame.Contents = translations[idx]
                if font:
                    try:
                        # 尝试设置字体（需确保 font 为 Illustrator 中有效的字体标识）
                        text_range = text_frame.TextRange
                        text_range.CharacterAttributes.TextFont = font
                    except Exception as e:
                        messagebox.showwarning("警告", f"设置字体失败: {str(e)}")
            elif mode == "add_below":
                new_text = doc.TextFrames.Add()
                new_text.Contents = translations[idx]
                new_text.Position = [text_frame.Position[0], text_frame.Position[1] - 20]
                if font:
                    try:
                        text_range = new_text.TextRange
                        text_range.CharacterAttributes.TextFont = font
                    except Exception as e:
                        messagebox.showwarning("警告", f"设置字体失败: {str(e)}")
        
        doc.Save()
        doc.Close()
        return True
    except Exception as e:
        messagebox.showerror("错误", f"更新失败: {str(e)}")
        return False

# --------------------------
# UI 界面
# --------------------------
class AIProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI 文件翻译处理工具")
        self.root.geometry("800x750")  # 增加高度以容纳新选项
        
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
                
        # 初始化变量（分别用于导出和导入）
        self.export_ai_file = ""
        self.export_ai_folder = ""
        self.export_csv_file = ""
        self.import_ai_file = ""
        self.import_csv_file = ""
        
        # 导出过滤选项
        self.export_numbers = tk.BooleanVar(value=True)  # 默认导出纯数字
        self.export_blanks = tk.BooleanVar(value=True)   # 默认导出空白内容
        self.merge_segments = tk.BooleanVar(value=False) # 默认不合并句段
        self.merge_threshold = tk.IntVar(value=1)       # 默认合并阈值
        
        # 创建 Notebook 选项卡
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 导出页：AI -> CSV
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="导出 AI → CSV")
        self.create_export_widgets(self.export_frame)
        
        # 导入页：CSV -> AI
        self.import_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.import_frame, text="导入 CSV → AI")
        self.create_import_widgets(self.import_frame)
        
        # 日志显示区域（共用）
        log_frame = ttk.Frame(root)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)
        ttk.Label(log_frame, text="处理日志:").pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state="disabled")
    
    # --------------------------
    # 导出页相关控件（增加合并选项）
    # --------------------------
    def create_export_widgets(self, frame):
        row = 0
        # AI 文件选择
        ttk.Label(frame, text="AI 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.entry_export_ai = ttk.Entry(frame, width=50)
        self.entry_export_ai.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览文件", command=self.browse_export_ai_file).grid(row=row, column=2, padx=5, pady=5)
        
        row += 1
        ttk.Label(frame, text="AI 文件夹路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.entry_export_folder = ttk.Entry(frame, width=50)
        self.entry_export_folder.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览文件夹", command=self.browse_export_ai_folder).grid(row=row, column=2, padx=5, pady=5)
        
        row += 1
        # CSV 文件路径
        ttk.Label(frame, text="CSV 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.entry_export_csv = ttk.Entry(frame, width=50)
        self.entry_export_csv.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览", command=self.browse_export_csv_file).grid(row=row, column=2, padx=5, pady=5)
        
        # 新增：句段合并选项
        row += 1
        merge_frame = ttk.LabelFrame(frame, text="句段合并设置")
        merge_frame.grid(row=row, column=0, columnspan=3, padx=10, pady=10, sticky="we")
        
        # 合并句段选项
        ttk.Checkbutton(merge_frame, text="合并相邻句段", variable=self.merge_segments).grid(
            row=0, column=0, padx=10, pady=5, sticky="w")
        
        # 合并阈值设置
        ttk.Label(merge_frame, text="合并阈值:").grid(row=0, column=1, padx=(20, 5), pady=5, sticky="e")
        self.spin_threshold = ttk.Spinbox(merge_frame, from_=1, to=200, width=5, textvariable=self.merge_threshold)
        self.spin_threshold.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Label(merge_frame, text="像素").grid(row=0, column=3, padx=(0, 10), pady=5, sticky="w")
        
        # 提示信息
        ttk.Label(merge_frame, text="* 合并垂直距离接近的文本框内容", foreground="gray").grid(
            row=1, column=0, columnspan=4, padx=10, pady=(0, 5), sticky="w")
        
        # 新增：导出内容过滤选项
        row += 1
        filter_frame = ttk.LabelFrame(frame, text="导出内容过滤")
        filter_frame.grid(row=row, column=0, columnspan=3, padx=10, pady=10, sticky="we")
        
        # 纯数字内容选项
        ttk.Checkbutton(filter_frame, text="导出纯数字内容", variable=self.export_numbers).grid(
            row=0, column=0, padx=10, pady=5, sticky="w")
        
        # 空白内容选项
        ttk.Checkbutton(filter_frame, text="导出空白内容", variable=self.export_blanks).grid(
            row=0, column=1, padx=10, pady=5, sticky="w")
        
        # 提示信息
        ttk.Label(filter_frame, text="* 取消勾选将跳过相应内容", foreground="gray").grid(
            row=1, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")
        
        row += 1
        # 导出按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="导出文字", command=self.export_text).pack(side=tk.LEFT, padx=5)
    
    def browse_export_ai_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("AI Files", "*.ai")])
        if file_path:
            self.export_ai_file = file_path
            self.entry_export_ai.delete(0, tk.END)
            self.entry_export_ai.insert(0, file_path)
            self.log("已选择 AI 文件: " + file_path)
            # 自动生成 CSV 文件路径
            default_csv = os.path.join(os.path.dirname(file_path), f"{os.path.splitext(os.path.basename(file_path))[0]}.csv")
            self.export_csv_file = default_csv
            self.entry_export_csv.delete(0, tk.END)
            self.entry_export_csv.insert(0, default_csv)
    
    def browse_export_ai_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.export_ai_folder = folder_path
            self.entry_export_folder.delete(0, tk.END)
            self.entry_export_folder.insert(0, folder_path)
            self.log("已选择 AI 文件夹: " + folder_path)
    
    def browse_export_csv_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )
        if file_path:
            self.export_csv_file = file_path
            self.entry_export_csv.delete(0, tk.END)
            self.entry_export_csv.insert(0, file_path)
            self.log("已设置 CSV 路径: " + file_path)
    
    def export_text(self):
        if not self.export_ai_file and not self.export_ai_folder:
            messagebox.showwarning("警告", "请先选择 AI 文件或文件夹")
            return
        
        # 获取导出选项状态
        export_numbers = self.export_numbers.get()
        export_blanks = self.export_blanks.get()
        merge_segments = self.merge_segments.get()
        merge_threshold = self.merge_threshold.get()
        
        # 记录设置
        settings_info = []
        if merge_segments:
            settings_info.append(f"合并句段(阈值={merge_threshold}px)")
        if not export_numbers:
            settings_info.append("跳过纯数字")
        if not export_blanks:
            settings_info.append("跳过空白内容")
        
        if settings_info:
            self.log(f"导出设置: {', '.join(settings_info)}")
        
        if self.export_ai_file:
            texts = extract_text_from_ai(
                self.export_ai_file, 
                merge_segments=merge_segments,
                merge_threshold=merge_threshold
            )
            
            if not texts:
                messagebox.showwarning("警告", "未提取到任何文本内容")
                return
                
            if not self.export_csv_file:
                self.export_csv_file = os.path.join(os.path.dirname(self.export_ai_file), 
                                                     f"{os.path.splitext(os.path.basename(self.export_ai_file))[0]}.csv")
            
            # 记录合并信息
            if merge_segments:
                self.log(f"合并后句段数量: {len(texts)}")
            
            generate_translation_csv(texts, self.export_csv_file, 
                                     os.path.basename(self.export_ai_file),
                                     export_numbers, export_blanks)
            self.log("成功导出翻译模板")
            
        elif self.export_ai_folder:
            output_folder = os.path.join(self.export_ai_folder, "output")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            processed_count = 0
            skipped_count = 0
            
            for filename in os.listdir(self.export_ai_folder):
                if filename.endswith(".ai"):
                    ai_file = os.path.join(self.export_ai_folder, filename)
                    
                    # 提取文本（应用合并设置）
                    texts = extract_text_from_ai(
                        ai_file, 
                        merge_segments=merge_segments,
                        merge_threshold=merge_threshold
                    )
                    
                    if texts:
                        output_csv = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.csv")
                        generate_translation_csv(texts, output_csv, filename, 
                                               export_numbers, export_blanks)
                        processed_count += 1
                        self.log(f"成功导出: {filename}")
                        
                        # 记录合并信息
                        if merge_segments:
                            self.log(f"  - 合并后句段数量: {len(texts)}")
                    else:
                        skipped_count += 1
                        self.log(f"跳过空文件: {filename}")
            
            self.log(f"批量导出完成: 处理 {processed_count} 个文件, 跳过 {skipped_count} 个文件")
    
    # --------------------------
    # 导入页相关控件
    # --------------------------
    def create_import_widgets(self, frame):
        row = 0
        # 选择需要更新的 AI 文件
        ttk.Label(frame, text="AI 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.entry_import_ai = ttk.Entry(frame, width=50)
        self.entry_import_ai.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览文件", command=self.browse_import_ai_file).grid(row=row, column=2, padx=5, pady=5)
        
        row += 1
        # 选择含有译文的 CSV 文件
        ttk.Label(frame, text="CSV 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.entry_import_csv = ttk.Entry(frame, width=50)
        self.entry_import_csv.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览文件", command=self.browse_import_csv_file).grid(row=row, column=2, padx=5, pady=5)
        
        row += 1
        # 预设翻译字体下拉选择框
        ttk.Label(frame, text="翻译字体:").grid(row=row, column=0, padx=5, pady=5, sticky="w")
        self.combo_font = ttk.Combobox(frame, values=["", "Helvetica", "Verdana", "Times New Roman", "Arial", "SimSun"],
                                       state="readonly", width=47)
        self.combo_font.grid(row=row, column=1, padx=5, pady=5)
        self.combo_font.set("")  # 默认为空
        
        row += 1
        # 操作按钮：替换译文 和 添加译文
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="替换译文", command=lambda: self.update_text_import("replace")).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="添加译文", command=lambda: self.update_text_import("add_below")).pack(side=tk.LEFT, padx=5)
    
    def browse_import_ai_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("AI Files", "*.ai")])
        if file_path:
            self.import_ai_file = file_path
            self.entry_import_ai.delete(0, tk.END)
            self.entry_import_ai.insert(0, file_path)
            self.log("已选择更新的 AI 文件: " + file_path)
    
    def browse_import_csv_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.import_csv_file = file_path
            self.entry_import_csv.delete(0, tk.END)
            self.entry_import_csv.insert(0, file_path)
            self.log("已选择译文 CSV 文件: " + file_path)
    
    def update_text_import(self, mode):
        if not self.import_ai_file:
            messagebox.showwarning("警告", "请先选择 AI 文件")
            return
        
        if not self.import_csv_file:
            messagebox.showwarning("警告", "请先选择 CSV 文件")
            return
        
        try:
            # 读取CSV文件
            translations = []
            with open(self.import_csv_file, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                next(reader)  # 跳过文件名行
                next(reader)  # 跳过标题行
                for row in reader:
                    if len(row) >= 2:  # 确保有译文列
                        translations.append(row[1])
            
            if not translations:
                messagebox.showwarning("警告", "CSV文件中未找到有效的译文")
                return
            
            font = self.combo_font.get().strip() or None
            if update_ai_file(self.import_ai_file, translations, mode, font):
                action = "替换" if mode == "replace" else "添加"
                self.log(f"成功{action}译文")
        except Exception as e:
            self.log(f"错误: {str(e)}")
            messagebox.showerror("错误", f"处理CSV文件失败: {str(e)}")
    
    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)

# --------------------------
# 启动程序
# --------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = AIProcessorApp(root)
    root.mainloop()