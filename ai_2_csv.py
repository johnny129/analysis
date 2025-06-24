import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv  # 添加csv模块
import comtypes.client  # 需要安装 comtypes

# --------------------------
# AI 文件处理核心逻辑
# --------------------------
def extract_text_from_ai(ai_file):
    """从 AI 文件中提取所有文本（需要 Adobe Illustrator 支持）"""
    try:
        ai = comtypes.client.CreateObject("Illustrator.Application")
        doc = ai.Open(ai_file)
        texts = []
        for text_frame in doc.TextFrames:
            texts.append(text_frame.Contents)
        doc.Close()
        return texts
    except Exception as e:
        messagebox.showerror("错误", f"无法提取文本: {str(e)}")
        return []

def generate_translation_csv(texts, output_csv, filename):
    """生成翻译用 CSV 文件"""
    try:
        with open(output_csv, 'w', encoding='utf-8-sig', newline='') as csvfile:  # 使用utf-8-sig编码
            writer = csv.writer(csvfile)
            # 写入文件名作为第一行（红色字体在CSV中不支持，但保留文件名标记）
            writer.writerow([f"文件名: {filename}"])
            writer.writerow(["原文", "译文"])  # 表头
            
            for text in texts:
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
        
        for idx, text_frame in enumerate(doc.TextFrames):
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
        self.root.geometry("800x600")
        
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap('icon.ico')
        except:
            pass
                
        # 初始化变量（分别用于导出和导入）
        self.export_ai_file = ""
        self.export_ai_folder = ""
        self.export_csv_file = ""  # 改为CSV文件
        self.import_ai_file = ""
        self.import_csv_file = ""  # 改为CSV文件
        
        # 创建 Notebook 选项卡
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 导出页：AI -> CSV
        self.export_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.export_frame, text="导出 AI → CSV")  # 修改标签
        self.create_export_widgets(self.export_frame)
        
        # 导入页：CSV -> AI
        self.import_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.import_frame, text="导入 CSV → AI")  # 修改标签
        self.create_import_widgets(self.import_frame)
        
        # 日志显示区域（共用）
        log_frame = ttk.Frame(root)
        log_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=5)
        ttk.Label(log_frame, text="处理日志:").pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state="disabled")
    
    # --------------------------
    # 导出页相关控件
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
        # CSV 文件路径（导出时若未手动设置，将自动沿用 AI 文件名）
        ttk.Label(frame, text="CSV 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")  # 修改标签
        self.entry_export_csv = ttk.Entry(frame, width=50)  # 改为CSV
        self.entry_export_csv.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览", command=self.browse_export_csv_file).grid(row=row, column=2, padx=5, pady=5)  # 修改命令
        
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
            # 自动生成 CSV 文件路径（保存在同目录下，文件名相同扩展名为 .csv）
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
    
    def browse_export_csv_file(self):  # 修改为CSV
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
        
        if self.export_ai_file:
            texts = extract_text_from_ai(self.export_ai_file)
            if not self.export_csv_file:
                self.export_csv_file = os.path.join(os.path.dirname(self.export_ai_file), 
                                                     f"{os.path.splitext(os.path.basename(self.export_ai_file))[0]}.csv")
            if texts:
                generate_translation_csv(texts, self.export_csv_file, os.path.basename(self.export_ai_file))
                self.log("成功导出翻译模板")
        elif self.export_ai_folder:
            output_folder = os.path.join(self.export_ai_folder, "output")
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
            
            for filename in os.listdir(self.export_ai_folder):
                if filename.endswith(".ai"):
                    ai_file = os.path.join(self.export_ai_folder, filename)
                    texts = extract_text_from_ai(ai_file)
                    if texts:
                        output_csv = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.csv")
                        generate_translation_csv(texts, output_csv, filename)
                        self.log(f"成功导出: {filename}")
            self.log("批量导出完成")
    
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
        ttk.Label(frame, text="CSV 文件路径:").grid(row=row, column=0, padx=5, pady=5, sticky="w")  # 修改标签
        self.entry_import_csv = ttk.Entry(frame, width=50)  # 改为CSV
        self.entry_import_csv.grid(row=row, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览文件", command=self.browse_import_csv_file).grid(row=row, column=2, padx=5, pady=5)  # 修改命令
        
        row += 1
        # 预设翻译字体下拉选择框，默认为空，不改变字体
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
    
    def browse_import_csv_file(self):  # 修改为CSV
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])  # 修改文件类型
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
            with open(self.import_csv_file, 'r', encoding='utf-8-sig') as csvfile:  # 使用相同的编码
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