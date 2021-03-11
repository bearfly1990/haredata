import webbrowser
import tkinter as tk
from PIL import ImageTk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
from app.common.tkinter import UIRoot
from app.feature.processor import DataProcessor
from app.common.log import logger
from app.common.io import PathUtils
from imgs.icon import Icon
import base64
import os
import traceback


class DataProcessorUI(UIRoot):
    entry_open_file = None

    def __init__(self):
        self.data_processor = DataProcessor()
        # super().__init__(title='test')
        self.root = Tk()
        self.root.resizable(width=False, height=False)
        self.root.title("数据处理小工具")

        # with open('tmp.ico', 'wb') as tmp:
        #     tmp.write(base64.b64decode(Icon().img))
        # self.root.iconbitmap('./tmp.ico')
        # os.remove('./tmp.ico')
        self.root.iconbitmap(PathUtils.resource_path('./imgs/logo.ico'))
        # self.root.iconbitmap('./imgs/login.ico')
        img_game = PhotoImage(file=PathUtils.resource_path('./imgs/game.png'))
        img_help = ImageTk.PhotoImage(file=PathUtils.resource_path("./imgs/help.png"))
        self.frame = self.create_new_frame()
        self.init_menu(img_game, img_help)
        # self.init_menu(None, None)
        self.init_base_info()
        self.init_control_buttons()
        tk.mainloop()

    def init_menu(self, img_game, img_help):
        menu_bar = Menu(self.root)
        self.root.config(menu=menu_bar)

        # 实例化菜单1，创建下拉菜单，调用add_separate创建分割线
        menu1 = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="编辑", menu=menu1)
        # menu1.add_command(label="Do Nothing")
        menu1.add_separator()
        menu1.add_command(label="退出", command=self.root.quit)

        menu2 = Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="更多", menu=menu2)
        # menu2.add_command(label="New Job", image=img_game, compound="left",
        #                   command=lambda: webbrowser.open("https://bearfly1990.github.io/"))
        #
        # menu2.add_command(label="Tkinter", image=img_help, compound="left", command=
        # lambda: webbrowser.open("https://bearfly1990.github.io/"))

    def open_file(self):
        self.entry_open_file.delete(0, tk.END)
        self.entry_open_file.insert(tk.END, filedialog.askopenfilename(title='选择Excel文件',
                                                                       filetypes=[('Excel', '*.xlsx'),
                                                                                  ('Excel', '*.xls'),
                                                                                  ('All Files', '*')]))

    def init_base_info(self):
        frame = self.create_new_frame(text='选择需要处理的文件')
        # tk.Label(frame, text='选择文件').grid(row=row, column=column, sticky=tk.W)
        button_open_file = tk.Button(frame, text="选择文件", command=self.open_file)
        # button_open_file.pack(fill=BOTH, expand=YES, padx=3, pady=3)
        button_open_file.grid(row=1, column=1, sticky=tk.W)
        self.entry_open_file = tk.Entry(frame, width=60)
        self.entry_open_file.grid(row=1, column=2)
        self.entry_open_file.insert(tk.END, '')

    def do_ok(self):
        try:
            self.root.withdraw()
            self.data_processor.field_input_file = self.entry_open_file.get()
            # self.data_processor.field_input_file = r'C:\Users\mayn\Desktop\权重计算测试基础数据.xlsx'
            logger.info(f'开始处理文件:{self.data_processor.field_input_file}')
            self.data_processor.process()
            logger.info('处理完成')
        except Exception as e:
            logger.error("运行出错，请确定选择了正确的文件和数据")
            logger.error(e)
            logger.error(traceback.format_exc())
        finally:
            self.root.deiconify()


def init_ui():
    DataProcessorUI()
