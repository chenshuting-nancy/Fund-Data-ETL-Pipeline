import tkinter as tk

class ModernButton(tk.Button):
    def __init__(self, master, **kwargs):
        kwargs.update({
            'relief': tk.FLAT,
            'bg': '#4A90E2',
            'fg': '#ffffff',
            'font': ('微软雅黑', 10),
            'cursor': 'hand2',
            'activebackground': '#357ABD',
            'activeforeground': '#ffffff',
            'padx': 15,
            'pady': 5,
            'bd': 0
        })
        super().__init__(master, **kwargs)
        self.bind('<Enter>', self.on_enter)
        self.bind('<Leave>', self.on_leave)

    def on_enter(self, e):
        self['background'] = '#357ABD'

    def on_leave(self, e):
        self['background'] = '#4A90E2'

class ModernEntry(tk.Entry):
    def __init__(self, master, **kwargs):
        kwargs.update({
            'relief': tk.FLAT,
            'bg': '#F5F5F7',
            'font': ('微软雅黑', 10),
            'bd': 0,
            'highlightthickness': 1,
            'highlightcolor': '#4A90E2',
            'highlightbackground': '#E0E0E0'
        })
        super().__init__(master, **kwargs) 