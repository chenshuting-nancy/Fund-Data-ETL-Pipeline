import tkinter as tk

def log(msg, log_text=None):
    """日志输出函数
    
    Args:
        msg: 要输出的消息
        log_text: tkinter.scrolledtext.ScrolledText对象，用于显示日志
    """
    print(msg)
    if log_text is not None:
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, msg + "\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)
        log_text.update() 