import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
from tkinter import *


class UIRoot(object):
    entry_user = None
    entry_pwd = None
    frame_control = None

    def __init__(self, title='UI'):
        # self.root = tk.Tk()
        # self.root.title(title)
        # self.frame = self.create_new_frame()
        pass

    def create_new_frame(self, text=None):
        frame = None
        if text:
            frame = ttk.Labelframe(self.root, text=text)
            frame.config(relief=tk.GROOVE)
        else:
            frame = tk.Frame(self.root)
            frame.config(relief=tk.GROOVE, bd=2)
        # frame = ttk.Labelframe(self.root, text=text) if text else ttk.Frame(self.root)
        frame.pack(fill=tk.X, expand=True, side=tk.TOP)

        return frame

    def init_input_field(self, frame=None, text='', val='', row=0, column=0, with_star=False, width=20):
        if frame is None:
            frame = self.frame
        tk.Label(frame, text=text).grid(row=row, column=column, sticky=tk.W)
        entry = tk.Entry(frame, width=width)
        entry.grid(row=row, column=column + 1)
        if with_star:
            entry.config(show="*")
        entry.insert(tk.END, val)
        return entry


    def init_user_pwd_entry(self, row=0, user_val=''):
        self.init_frame()
        self.entry_user = self.init_input_field(text="Username:", val=user_val, row=row, column=0)
        self.entry_pwd = self.init_input_field(text="password:", row=row + 1, column=0, with_star=True)


    def init_control_buttons(self, row=0):
        self.frame_control = self.create_new_frame()
        tk.Button(self.frame_control, text='开始', command=self.do_ok).grid(row=row, column=0, sticky=tk.W, pady=4)
        tk.Button(self.frame_control, text='退出', command=self.frame.quit).grid(row=row, column=1, sticky=tk.W, pady=4)
