import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
import subprocess
import time
from ui.modern_widgets import ModernButton, ModernEntry
from ui.product_code_manager import ProductCodeManager
from extractors.dividend_extractor import run_dividend_extract
from extractors.purchase_extractor import run_purchase_extract
from extractors.purchase_confirm_extractor import run_purchase_confirm_extract
from extractors.redemption_extractor import run_redemption_extract
from extractors.conversion_extractor import run_conversion_extract
from extractors.manual_purchase_apply_extractor import run_manual_purchase_apply_extract
from extractors.manual_redemption_extractor import run_manual_redemption_extract
from extractors.manual_purchase_confirm_extractor import run_manual_purchase_confirm_extract
from extractors.manual_dividen_extractor import run_manual_dividend_extract

DEFAULT_JSON_FILENAME = "product_codes.json"
DEFAULT_CONVERSION_JSON_FILENAME = "product_codes_conversion.json"  # 新增
DEFAULT_FOLDER_PATH = r"\\172.18.101.248\估值邮件\估值材料（备查）"

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("基金单提取程序")
        self.geometry("900x750")  # 增加高度以容纳新的输入框
        self.configure(bg='#ffffff')
        
        try:
            self.attributes('-transparent', True)
            self.attributes('-alpha', 0.98)
        except:
            pass

        # 默认路径设置
        self.folder_path = DEFAULT_FOLDER_PATH
        self.json_path = os.path.join(self.folder_path, DEFAULT_JSON_FILENAME)
        self.conversion_json_path = os.path.join(self.folder_path, DEFAULT_CONVERSION_JSON_FILENAME)  # 新增
        self.status = None  # 用于记录最近一次输出目录
        self.is_extracting = False  # 添加标志位防止重复执行
        self.create_widgets()
        self.center_window()

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 900) // 2
        y = (screen_height - 750) // 2  # 调整高度
        self.geometry(f'900x750+{x}+{y}')

    def create_widgets(self):
        # 主容器
        main_container = tk.Frame(self, bg='#ffffff')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        tk.Label(
            main_container,
            text="基金单提取程序",
            font=('微软雅黑', 16, 'bold'),
            bg='#ffffff',
            fg='#000000'
        ).pack(pady=(0, 20))

        # 替换 Text 组件为 Label 组件
        platform_frame = tk.Frame(main_container, bg='#ffffff')
        platform_frame.pack(fill=tk.X, pady=(0, 20))
        
        # 主标题
        tk.Label(
            platform_frame,
            text="支持的平台：",
            font=('微软雅黑', 11, 'bold'),
            bg='#ffffff',
            fg='#333333',
            anchor='w'
        ).pack(fill=tk.X, anchor='w')
        
        # 共同支持
        content_frame_all = tk.Frame(platform_frame, bg='#ffffff')
        content_frame_all.pack(fill=tk.X, pady=(5, 0), anchor='w')
        
        tk.Label(
            content_frame_all,
            text="• 分红、申购、赎回均支持：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')
        
        tk.Label(
            content_frame_all,
            text="好买、天天、利得、盈米、长量、平安行 E通、交行交e通、网金、腾元、和讯、京东、融联创、联泰、民生同业e+、攀赢基金、证达通",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700  # 自动换行
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        # 分红单
        content_frame = tk.Frame(platform_frame, bg='#ffffff')
        content_frame.pack(fill=tk.X, pady=(5, 0), anchor='w')
        
        tk.Label(
            content_frame,
            text="• 分红单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')
        
        tk.Label(
            content_frame,
            text="兴证、招银、邮储、建行直销、基煜基金、宁波银行、国信嘉利基金",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700  # 自动换行
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))
        
        # 申购申请单
        content_frame2 = tk.Frame(platform_frame, bg='#ffffff')
        content_frame2.pack(fill=tk.X, pady=(5, 0), anchor='w')
        
        tk.Label(
            content_frame2,
            text="• 申购申请单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')
        
        tk.Label(
            content_frame2,
            text="招银、基煜基金、宁波银行、国信嘉利基金",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))
        
        # 申购确认单
        content_frame3 = tk.Frame(platform_frame, bg='#ffffff')
        content_frame3.pack(fill=tk.X, pady=(5, 0), anchor='w')
        
        tk.Label(
            content_frame3,
            text="• 申购确认单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')
        
        tk.Label(
            content_frame3,
            text="兴证、招银、邮储、建行直销、基煜基金、宁波银行、国信嘉利基金",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        # 赎回确认单
        content_frame4 = tk.Frame(platform_frame, bg='#ffffff')
        content_frame4.pack(fill=tk.X, pady=(5, 0), anchor='w')
        
        tk.Label(
            content_frame4,
            text="• 赎回确认单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')
        
        tk.Label(
            content_frame4,
            text="建行直销、京东超级转换强行赎回",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        tk.Label(
            content_frame4,
            text="请检查赎回到账日期",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='red',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        # 超级转换确认单
        content_frame5 = tk.Frame(platform_frame, bg='#ffffff')
        content_frame5.pack(fill=tk.X, pady=(5, 0), anchor='w')

        tk.Label(
            content_frame5,
            text="• 超级转换确认单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')

        tk.Label(
            content_frame5,
            text="京东肯特瑞、天天基金",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        # 理财产品手工单
        content_frame6 = tk.Frame(platform_frame, bg='#ffffff')
        content_frame6.pack(fill=tk.X, pady=(5, 0), anchor='w')

        tk.Label(
            content_frame6,
            text="• 理财产品手工单：",
            font=('微软雅黑', 10, 'bold'),
            bg='#ffffff',
            fg='#333333'
        ).pack(side=tk.LEFT, anchor='w')

        tk.Label(
            content_frame6,
            text="申购申请单、申购确认单、赎回确认单",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='#666666',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))
        # 新增红色提示
        tk.Label(
            content_frame6,
            text="申购返款、赎回到账、红利到账需要手动录入",
            font=('微软雅黑', 9),
            bg='#ffffff',
            fg='red',
            wraplength=700
        ).pack(side=tk.LEFT, anchor='w', padx=(5, 0))

        # 文件选择区域
        file_frame = tk.Frame(main_container, bg='#ffffff')
        file_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            file_frame,
            text="主目录路径：",
            font=('微软雅黑', 10),
            bg='#ffffff'
        ).pack(side=tk.LEFT)

        self.path_entry = ModernEntry(file_frame)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 10))
        self.path_entry.insert(0, self.folder_path)

        ModernButton(
            file_frame,
            text="选择主目录",
            command=self.select_folder
        ).pack(side=tk.LEFT)

        # JSON文件选择区域
        json_frame = tk.Frame(main_container, bg='#ffffff')
        json_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            json_frame,
            text="映射文件路径：",
            font=('微软雅黑', 10),
            bg='#ffffff'
        ).pack(side=tk.LEFT)

        self.json_entry = ModernEntry(json_frame)
        self.json_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 10))
        self.json_entry.insert(0, self.json_path)

        ModernButton(
            json_frame,
            text="选择文件",
            command=self.select_json
        ).pack(side=tk.LEFT)

        ModernButton(
            json_frame,
            text="管理映射",
            command=self.manage_codes
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 转换单JSON文件选择区域（新增）
        conversion_json_frame = tk.Frame(main_container, bg='#ffffff')
        conversion_json_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            conversion_json_frame,
            text="转换单映射文件：",
            font=('微软雅黑', 10),
            bg='#ffffff'
        ).pack(side=tk.LEFT)

        self.conversion_json_entry = ModernEntry(conversion_json_frame)
        self.conversion_json_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 10))
        self.conversion_json_entry.insert(0, self.conversion_json_path)

        ModernButton(
            conversion_json_frame,
            text="选择文件",
            command=self.select_conversion_json
        ).pack(side=tk.LEFT)

        ModernButton(
            conversion_json_frame,
            text="管理映射",
            command=self.manage_conversion_codes
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 一键提取按钮区域（新增）
        one_click_frame = tk.Frame(main_container, bg='#ffffff')
        one_click_frame.pack(fill=tk.X, pady=(10, 5))
        
        self.one_click_btn = ModernButton(
            one_click_frame,
            text="一键提取所有单据",
            command=self.extract_all
        )
        self.one_click_btn.pack(side=tk.LEFT, padx=5)

        ModernButton(
            one_click_frame,
            text="打开输出文件夹",
            command=self.open_output_folder
        ).pack(side=tk.LEFT, padx=5)  # 放在一键提取按钮旁边
        
        # 添加显示/隐藏复选框
        self.show_advanced = tk.BooleanVar(value=False)
        tk.Checkbutton(
            one_click_frame,
            text="显示单独提取选项",
            variable=self.show_advanced,
            command=self.toggle_advanced,
            bg='#ffffff',
            font=('微软雅黑', 9),
            activebackground='#ffffff'
        ).pack(side=tk.RIGHT, padx=10)

        # 创建一个专门的容器来放置advanced_frame（新增这部分）
        self.advanced_container = tk.Frame(main_container, bg='#ffffff')
        self.advanced_container.pack(fill=tk.X, pady=0)

        # 创建一个容器来包含所有单独的按钮
        self.advanced_frame = tk.Frame(self.advanced_container, bg='#ffffff')  # 注意父容器改为advanced_container

        # 按钮区域 - 改为放在 advanced_frame 中
        btn_frame = tk.Frame(self.advanced_frame, bg='#ffffff')  # 注意改为 self.advanced_frame
        btn_frame.pack(fill=tk.X, pady=10)

        ModernButton(
            btn_frame,
            text="开始提取分红单",
            command=lambda: self.start_extract('dividend')
        ).pack(side=tk.LEFT, padx=5)

        ModernButton(
            btn_frame,
            text="开始提取申购申请单",
            command=lambda: self.start_extract('purchase')
        ).pack(side=tk.LEFT, padx=5)

        ModernButton(
            btn_frame,
            text="开始提取申购确认单",
            command=lambda: self.start_extract('purchase_confirm')
        ).pack(side=tk.LEFT, padx=5)

        ModernButton(
            btn_frame,
            text="开始提取赎回确认单",
            command=lambda: self.start_extract('redemption')
        ).pack(side=tk.LEFT, padx=5)

        ModernButton(
            btn_frame,
            text="开始提取超级转换确认单",
            command=lambda: self.start_extract('conversion')
        ).pack(side=tk.LEFT, padx=5)

        # 新增一行按钮区域 - 也改为放在 advanced_frame 中
        btn_frame2 = tk.Frame(self.advanced_frame, bg='#ffffff')  # 注意改为 self.advanced_frame
        btn_frame2.pack(fill=tk.X, pady=0)
        ModernButton(
            btn_frame2,
            text="开始提取万事如意申购申请单",
            command=lambda: self.start_extract('manual_purchase_apply')
        ).pack(side=tk.LEFT, padx=5)
        ModernButton(
            btn_frame2,
            text="开始提取万事如意申购确认单",
            command=lambda: self.start_extract('manual_purchase_confirm')
        ).pack(side=tk.LEFT, padx=5)
        ModernButton(
            btn_frame2,
            text="开始提取万事如意赎回确认单",
            command=lambda: self.start_extract('manual_redemption')
        ).pack(side=tk.LEFT, padx=5)
        ModernButton(
            btn_frame2,
            text="开始提取万事如意红利除权单",
            command=lambda: self.start_extract('manual_dividend')
        ).pack(side=tk.LEFT, padx=5)


        # 日志区域
        log_frame = tk.Frame(main_container, bg='#ffffff')
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            font=('微软雅黑', 10),
            bg='#F5F5F7',
            fg='#000000'
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_path)

    def select_json(self):
        json_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if json_path:
            self.json_path = json_path
            self.json_entry.delete(0, tk.END)
            self.json_entry.insert(0, json_path)

    def select_conversion_json(self):  # 新增方法
        json_path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if json_path:
            self.conversion_json_path = json_path
            self.conversion_json_entry.delete(0, tk.END)
            self.conversion_json_entry.insert(0, json_path)

    def manage_codes(self):
        ProductCodeManager(self, self.json_path)

    def manage_conversion_codes(self):  # 新增方法
        ProductCodeManager(self, self.conversion_json_path, title="转换单产品代码映射管理")

    def toggle_advanced(self):
        """切换高级选项的显示/隐藏"""
        if self.show_advanced.get():
            # 显示容器（容器会自动带着advanced_frame一起显示）
            self.advanced_container.pack(fill=tk.X, pady=(5, 10), after=self.one_click_btn.master)
            self.advanced_frame.pack(fill=tk.X)
            self.geometry("900x850")
        else:
            # 隐藏整个容器
            self.advanced_container.pack_forget()
            self.geometry("900x750")

    def extract_all(self):
        """一键提取所有单据"""
        if self.is_extracting:
            messagebox.showinfo("提示", "正在执行提取任务，请稍候...")
            return
        
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择文件夹！")
            return
        
        # 检查映射文件
        if not os.path.exists(self.json_path):
            messagebox.showwarning("警告", "映射文件不存在！")
            return
        
        if not os.path.exists(self.conversion_json_path):
            messagebox.showwarning("警告", "转换单映射文件不存在！")
            return
        
        # 清空日志
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.insert(tk.END, "========== 开始一键提取所有单据 ==========\n\n")
        self.log_text.config(state=tk.DISABLED)
        
        def task():
            self.is_extracting = True
            self.one_click_btn.config(state=tk.DISABLED, text="正在提取中...")
            
            # 定义提取顺序和对应的任务
            extract_tasks = [
                ('dividend', '分红单', run_dividend_extract, self.json_path),
                ('purchase', '申购申请单', run_purchase_extract, self.json_path),
                ('purchase_confirm', '申购确认单', run_purchase_confirm_extract, self.json_path),
                ('redemption', '赎回确认单', run_redemption_extract, self.json_path),
                ('conversion', '超级转换确认单', run_conversion_extract, self.conversion_json_path),
                ('manual_purchase_apply', '万事如意申购申请单', run_manual_purchase_apply_extract, self.json_path),
                ('manual_purchase_confirm', '万事如意申购确认单', run_manual_purchase_confirm_extract, self.json_path),
                ('manual_redemption', '万事如意赎回确认单', run_manual_redemption_extract, self.json_path),
                ('manual_dividend', '万事如意红利除权单', run_manual_dividend_extract, self.json_path)
            ]
            
            try:
                for task_type, task_name, extract_func, json_path in extract_tasks:
                    # 在日志中添加分隔线
                    self.log_text.config(state=tk.NORMAL)
                    self.log_text.insert(tk.END, f"\n---------- 正在提取{task_name} ----------\n")
                    self.log_text.see(tk.END)
                    self.log_text.config(state=tk.DISABLED)
                    
                    # 执行提取任务
                    result = extract_func(self.folder_path, json_path, self.log_text)
                    if result:
                        self.status = result
                    
                    # 每个任务之间稍微延迟，避免过快
                    time.sleep(0.5)
                
                # 完成所有任务
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, "\n========== 所有单据提取完成！ ==========\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                
                messagebox.showinfo("完成", "所有单据提取完成！")
                
            except Exception as e:
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, f"\n错误：{str(e)}\n")
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
                messagebox.showerror("错误", f"提取过程中出现错误：{str(e)}")
            finally:
                self.is_extracting = False
                self.one_click_btn.config(state=tk.NORMAL, text="一键提取所有单据")
        
        # 在新线程中执行任务
        threading.Thread(target=task, daemon=True).start()

    def start_extract(self, extract_type):
        if not self.folder_path:
            messagebox.showwarning("警告", "请先选择文件夹！")
            return
        
        # 根据提取类型检查对应的映射文件
        if extract_type == 'conversion':
            if not os.path.exists(self.conversion_json_path):
                messagebox.showwarning("警告", "转换单映射文件不存在！")
                return
        else:
            if not os.path.exists(self.json_path):
                messagebox.showwarning("警告", "映射文件不存在！")
                return

        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        def task():
            try:
                result = None
                if extract_type == 'dividend':
                    result = run_dividend_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'purchase':
                    result = run_purchase_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'purchase_confirm':
                    result = run_purchase_confirm_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'redemption':
                    result = run_redemption_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'conversion':
                    # 传递转换单专用的映射文件路径
                    result = run_conversion_extract(self.folder_path, self.conversion_json_path, self.log_text)
                elif extract_type == 'manual_purchase_apply':
                    result = run_manual_purchase_apply_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'manual_purchase_confirm':
                    result = run_manual_purchase_confirm_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'manual_redemption':
                    result = run_manual_redemption_extract(self.folder_path, self.json_path, self.log_text)
                elif extract_type == 'manual_dividend':
                    result = run_manual_dividend_extract(self.folder_path, self.json_path, self.log_text)
                if result:
                    self.status = result
            except Exception as e:
                messagebox.showerror("错误", str(e))

        threading.Thread(target=task, daemon=True).start()

    def open_output_folder(self):
        if self.status and os.path.isdir(self.status):
            try:
                if os.name == "nt":
                    os.startfile(self.status)
                else:
                    subprocess.call(["open", self.status])
            except Exception as e:
                messagebox.showerror("错误", f"无法打开输出文件夹：{str(e)}")
        else:
            messagebox.showwarning("警告", "没有可用的输出文件夹！")

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()