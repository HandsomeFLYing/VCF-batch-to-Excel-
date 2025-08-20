import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import threading

class VcfPhoneExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("VCF文件电话号码提取器")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置字体，确保中文显示正常
        self.font = ('SimHei', 10)
        
        # 文件夹路径变量
        self.folder_path = tk.StringVar()
        
        # 输出文件路径变量
        self.output_path = tk.StringVar(value=os.path.join(os.getcwd(), "vcf电话号码列表.xlsx"))
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部框架：选择文件夹
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)
        
        tk.Label(top_frame, text="VCF文件夹路径：", font=self.font).pack(side=tk.LEFT, padx=5)
        
        path_entry = tk.Entry(top_frame, textvariable=self.folder_path, width=50, font=self.font)
        path_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        browse_btn = tk.Button(top_frame, text="浏览...", command=self.browse_folder, font=self.font)
        browse_btn.pack(side=tk.LEFT, padx=5)
        
        # 中间框架：输出文件设置
        mid_frame = tk.Frame(self.root, padx=10, pady=5)
        mid_frame.pack(fill=tk.X)
        
        tk.Label(mid_frame, text="输出文件路径：", font=self.font).pack(side=tk.LEFT, padx=5)
        
        output_entry = tk.Entry(mid_frame, textvariable=self.output_path, width=50, font=self.font)
        output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        output_btn = tk.Button(mid_frame, text="选择...", command=self.browse_output, font=self.font)
        output_btn.pack(side=tk.LEFT, padx=5)
        
        # 格式选择框架
        format_frame = tk.Frame(self.root, padx=10, pady=5)
        format_frame.pack(fill=tk.X)
        
        self.format_var = tk.StringVar(value="xlsx")
        tk.Label(format_frame, text="输出格式：", font=self.font).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(format_frame, text="Excel (.xlsx)", variable=self.format_var, value="xlsx", font=self.font).pack(side=tk.LEFT, padx=5)
        tk.Radiobutton(format_frame, text="CSV (.csv)", variable=self.format_var, value="csv", font=self.font).pack(side=tk.LEFT, padx=5)
        
        # 按钮框架
        btn_frame = tk.Frame(self.root, padx=10, pady=10)
        btn_frame.pack()
        
        self.generate_btn = tk.Button(btn_frame, text="提取电话号码", command=self.start_extraction, font=('SimHei', 12, 'bold'), 
                                      width=15, height=1, bg="#4CAF50", fg="white")
        self.generate_btn.pack()
        
        # 日志区域
        log_frame = tk.Frame(self.root, padx=10, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(log_frame, text="处理日志：", font=self.font).pack(anchor=tk.W)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=self.font, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text.config(state=tk.DISABLED)
        
    def browse_folder(self):
        """选择VCF文件所在文件夹"""
        folder = filedialog.askdirectory(title="选择VCF文件所在文件夹")
        if folder:
            self.folder_path.set(folder)
            # 自动建议输出路径
            default_output = os.path.join(folder, "vcf电话号码列表." + self.format_var.get())
            self.output_path.set(default_output)
            self.log(f"已选择文件夹：{folder}")
            
    def browse_output(self):
        """选择输出文件路径"""
        file_ext = ".xlsx" if self.format_var.get() == "xlsx" else ".csv"
        file_type = f"Excel文件 (*{file_ext})" if self.format_var.get() == "xlsx" else f"CSV文件 (*{file_ext})"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=file_ext,
            filetypes=[(file_type, f"*{file_ext}"), ("所有文件", "*.*")],
            title="保存输出文件"
        )
        if file_path:
            self.output_path.set(file_path)
            self.log(f"输出文件将保存为：{file_path}")
            
    def log(self, message):
        """在日志区域显示消息"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)  # 滚动到最后一行
        self.log_text.config(state=tk.DISABLED)
        
    def start_extraction(self):
        """开始提取电话号码（在新线程中运行以避免界面冻结）"""
        folder = self.folder_path.get()
        output = self.output_path.get()
        
        if not folder:
            messagebox.showerror("错误", "请先选择VCF文件所在文件夹")
            return
            
        if not output:
            messagebox.showerror("错误", "请指定输出文件路径")
            return
            
        # 禁用生成按钮，防止重复点击
        self.generate_btn.config(state=tk.DISABLED, text="处理中...")
        self.log("开始提取VCF文件中的电话号码...")
        
        # 在新线程中执行处理，避免界面冻结
        threading.Thread(target=self.extract_phones, args=(folder, output), daemon=True).start()
        
    def extract_phones(self, folder_path, output_path):
        """提取VCF文件中的电话号码"""
        try:
            # 检查文件夹是否存在
            if not os.path.exists(folder_path):
                self.log(f"错误：文件夹不存在 - {folder_path}")
                messagebox.showerror("错误", f"文件夹不存在：{folder_path}")
                return
                
            # 收集VCF文件信息和电话号码
            vcf_info = []
            file_count = 0
            total_phones = 0
            
            # 遍历文件夹
            for filename in os.listdir(folder_path):
                if filename.lower().endswith('.vcf'):
                    file_count += 1
                    self.log(f"正在处理：{filename}")
                    
                    file_path = os.path.join(folder_path, filename)
                    try:
                        # 尝试不同编码读取文件，提高兼容性
                        encodings = ['utf-8', 'gbk', 'latin-1']
                        content = None
                        
                        for encoding in encodings:
                            try:
                                with open(file_path, 'r', encoding=encoding) as f:
                                    content = f.read()
                                break
                            except UnicodeDecodeError:
                                continue
                                
                        if content is None:
                            content = "无法解析文件内容（编码问题）"
                            phones = ["文件解析失败"]
                        else:
                            # 提取电话号码 - 匹配TEL开头的行
                            # 正则表达式：匹配以TEL开头，然后是可能的参数，最后是电话号码
                            phone_pattern = r'TEL(;.*?)?:(.*?)(\r?\n|$)'
                            matches = re.findall(phone_pattern, content, re.IGNORECASE)
                            
                            # 提取并清理电话号码
                            phones = []
                            for match in matches:
                                phone = match[1].strip()
                                # 移除可能的非数字字符（保留+号，用于国际号码）
                                cleaned_phone = re.sub(r'[^\d+]', '', phone)
                                if cleaned_phone:
                                    phones.append(cleaned_phone)
                            
                            total_phones += len(phones)
                            
                            if not phones:
                                phones = ["未找到电话号码"]
                        
                        vcf_info.append({
                            '文件名': filename,
                            '电话号码': '; '.join(phones),  # 多个号码用分号分隔
                            '文件内容': content
                        })
                    except Exception as e:
                        self.log(f"处理文件 {filename} 时出错：{str(e)}")
            
            if not vcf_info:
                self.log("未找到任何VCF文件")
                messagebox.showinfo("提示", "未找到任何VCF文件")
                return
                
            # 创建DataFrame并保存
            df = pd.DataFrame(vcf_info)
            
            if self.format_var.get() == "xlsx":
                df.to_excel(output_path, index=False, engine='openpyxl')
            else:
                df.to_csv(output_path, index=False, encoding='utf-8-sig')
                
            self.log(f"处理完成！共找到 {file_count} 个VCF文件")
            self.log(f"共提取到 {total_phones} 个电话号码")
            self.log(f"文件已保存至：{output_path}")
            messagebox.showinfo("成功", 
                              f"已完成提取，共处理 {file_count} 个VCF文件\n"
                              f"提取到 {total_phones} 个电话号码\n"
                              f"保存路径：{output_path}")
            
        except Exception as e:
            error_msg = f"处理过程出错：{str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
        finally:
            # 恢复按钮状态
            self.root.after(0, lambda: self.generate_btn.config(state=tk.NORMAL, text="提取电话号码"))

if __name__ == "__main__":
    root = tk.Tk()
    app = VcfPhoneExtractor(root)
    root.mainloop()
    