import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from .modern_widgets import ModernButton

class ProductCodeManager(tk.Toplevel):
    def __init__(self, master, json_path, title="产品代码映射管理", *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.window_title = title  # 保存标题
        self.title(title)
        self.geometry("600x500")
        self.configure(bg='#ffffff')
        
        try:
            self.attributes('-transparent', True)
            self.attributes('-alpha', 0.98)
        except:
            pass

        self.json_path = json_path
        self.code_dict = {}
        # 添加排序状态变量
        self.sort_column = "账套编号"  # 默认按账套编号排序
        self.sort_reverse = False  # False为升序，True为降序
        self.create_widgets()
        self.load_json()
        self.center_window()

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 600) // 2
        y = (screen_height - 500) // 2
        self.geometry(f'600x500+{x}+{y}')

    def treeview_sort_column(self, col):
        """列排序函数"""
        if self.sort_column == col:
            # 如果点击的是同一列，则反转排序顺序
            self.sort_reverse = not self.sort_reverse
        else:
            # 如果点击的是不同列，则设置新的排序列，并默认升序
            self.sort_column = col
            self.sort_reverse = False
        
        # 获取所有项目
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        
        # 根据列类型进行排序
        if col == "账套编号":
            # 账套编号按数字排序，如果不是数字则按字符串排序
            def sort_key(x):
                try:
                    return (0, int(x[0]))  # 数字排在前面
                except ValueError:
                    return (1, x[0])  # 字符串排在后面
            
            l.sort(key=sort_key, reverse=self.sort_reverse)
        else:
            # 其他列按字符串排序
            l.sort(reverse=self.sort_reverse)
        
        # 重新排列项目
        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)
        
        # 更新列标题显示排序方向
        for col_name in ["产品名称", "账套编号"]:
            if col_name == col:
                self.tree.heading(col_name, text=f"{col_name} {'↓' if self.sort_reverse else '↑'}")
            else:
                self.tree.heading(col_name, text=col_name)

    def create_widgets(self):
        main_container = tk.Frame(self, bg='#ffffff')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # 标题
        tk.Label(
            main_container,
            text=self.window_title, # 使用保存的标题
            font=('微软雅黑', 14, 'bold'),
            bg='#ffffff',
            fg='#000000'
        ).pack(pady=(0, 20))

        # 树形视图
        style = ttk.Style()
        style.configure(
            "Treeview",
            font=('微软雅黑', 12),
            rowheight=30,
            background='#F5F5F7',
            fieldbackground='#F5F5F7'
        )
        style.configure("Treeview.Heading", font=('微软雅黑', 12))

        self.tree = ttk.Treeview(
            main_container,
            columns=("产品名称", "账套编号"),
            show="headings"
        )
        # 设置列标题和点击事件
        for col in ("产品名称", "账套编号"):
            self.tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(c))
        self.tree.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 按钮区域
        btn_frame = tk.Frame(main_container, bg='#ffffff')
        btn_frame.pack(fill=tk.X, pady=10)

        ModernButton(btn_frame, text="新增", command=self.add_entry).pack(side=tk.LEFT, padx=2)
        ModernButton(btn_frame, text="修改", command=self.edit_entry).pack(side=tk.LEFT, padx=2)
        ModernButton(btn_frame, text="删除", command=self.delete_entry).pack(side=tk.LEFT, padx=2)
        ModernButton(btn_frame, text="保存", command=self.save_json).pack(side=tk.RIGHT, padx=2)
        ModernButton(btn_frame, text="刷新", command=self.load_json).pack(side=tk.RIGHT, padx=2)

    def load_json(self):
        self.code_dict = {}
        try:
            if os.path.exists(self.json_path):
                with open(self.json_path, "r", encoding="utf-8") as f:
                    self.code_dict = json.load(f)
        except Exception as e:
            messagebox.showerror("错误", f"加载JSON失败: {e}")
            return
        self.refresh_view()

    def refresh_view(self):
        self.tree.delete(*self.tree.get_children())
        items = list(self.code_dict.items())
        
        # 根据当前排序设置对列表进行排序
        if self.sort_column == "账套编号":
            # 尝试按数字排序，如果失败则按字符串排序
            def sort_key(x):
                try:
                    return (0, int(x[1]))  # 数字排在前面
                except ValueError:
                    return (1, x[1])  # 字符串排在后面
            
            items.sort(key=sort_key, reverse=self.sort_reverse)
        else:  # 产品名称
            items.sort(key=lambda x: x[0], reverse=self.sort_reverse)
        
        # 插入排序后的项目
        for k, v in items:
            self.tree.insert("", tk.END, values=(k, v))
        
        # 更新列标题显示排序方向
        for col in ["产品名称", "账套编号"]:
            if col == self.sort_column:
                self.tree.heading(col, text=f"{col} {'↓' if self.sort_reverse else '↑'}")
            else:
                self.tree.heading(col, text=col)

    def add_entry(self):
        pname = simpledialog.askstring("新增", "产品名称：", parent=self)
        if not pname:
            return
        code = simpledialog.askstring("新增", "账套编号：", parent=self)
        if not code:
            return
        
        # 尝试转换为整数，如果失败则保持为字符串
        try:
            code = int(code)
        except ValueError:
            pass  # 保持为字符串
        
        if pname in self.code_dict:
            messagebox.showerror("错误", "该产品已存在！")
            return
        self.code_dict[pname] = code
        self.refresh_view()

    def edit_entry(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请选择要修改的条目")
            return
        item = self.tree.item(sel[0])
        pname, code = item["values"]
        new_pname = simpledialog.askstring("修改产品名称", "产品名称：", initialvalue=pname, parent=self)
        if not new_pname:
            return
        new_code = simpledialog.askstring("修改账套编号", "账套编号：", initialvalue=str(code), parent=self)
        if not new_code:
            return
        
        # 尝试转换为整数，如果失败则保持为字符串
        try:
            new_code = int(new_code)
        except ValueError:
            pass  # 保持为字符串
        
        # 处理重命名
        if new_pname != pname and new_pname in self.code_dict:
            messagebox.showerror("错误", "新产品名称已存在！")
            return
        del self.code_dict[pname]
        self.code_dict[new_pname] = new_code
        self.refresh_view()

    def delete_entry(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请选择要删除的条目")
            return
        item = self.tree.item(sel[0])
        pname = item["values"][0]
        if messagebox.askyesno("确认", f"确定要删除产品：{pname}？"):
            del self.code_dict[pname]
            self.refresh_view()
            # 自动保存更改
            try:
                with open(self.json_path, "w", encoding="utf-8") as f:
                    json.dump(self.code_dict, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("删除成功", f"已删除产品：{pname}")
            except Exception as e:
                messagebox.showerror("保存失败", f"删除成功但保存文件失败: {str(e)}")

    def save_json(self):
        try:
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(self.code_dict, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存成功", "产品代码映射已保存！")
        except Exception as e:
            messagebox.showerror("保存失败", str(e)) 