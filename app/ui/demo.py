import webbrowser
import tkinter as tk
from PIL import ImageTk
from tkinter import ttk
from tkinter import messagebox
from tkinter import *
from app.common.tkinter import UIRoot


class Survey():
    field_user = None
    field_age = None
    field_sex = None
    field_vips = None

    def __init__(self):
        pass


class SurveyUI(UIRoot):
    def __init__(self):
        self.survey = Survey()
        # super().__init__(title='test')
        self.root = Tk()
        self.root.resizable(width=False, height=False)
        self.root.title("调查表")
        self.root.iconbitmap('./imgs/login.ico')
        img_game = PhotoImage(file='./imgs/game.png')
        img_help = ImageTk.PhotoImage(file="./imgs/help.png")
        self.frame = self.create_new_frame()
        self.init_menu(img_game, img_help)
        self.init_base_info()
        self.init_q2()
        self.init_q3()
        self.init_control_buttons()
        tk.mainloop()

    def init_menu(self, img_game, img_help):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        # 实例化菜单1，创建下拉菜单，调用add_separate创建分割线
        menu1 = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="编辑", menu=menu1)
        menu1.add_command(label="Do Nothing")
        menu1.add_separator()
        menu1.add_command(label="退出", command=self.root.quit)

        menu2 = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="更多", menu=menu2)
        menu2.add_command(label="New Job", image=img_game, compound="left", command=lambda: gamerun())

        menu2.add_command(label="Tkinter", image=img_help, compound="left", command=
        lambda: webbrowser.open("http://effbot.org/tkinterbook/tkinter-index.htm"))

    def init_base_info(self):
        frame = self.create_new_frame()
        self.survey.field_user = self.init_input_field(frame, text='姓名:',
                                                       val='', row=0, column=0, with_star=False, width=20)
        frame = self.create_new_frame()
        self.survey.field_age = self.init_input_field(frame, text='年龄:',
                                                      val='', row=1, column=0, with_star=False, width=5)
        frame = self.create_new_frame()
        # frame_radio = ttk.Labelframe(frame, text='', padding=20)
        # frame_radio.pack(fill=BOTH, expand=YES, padx=10, pady=10)
        Label(frame, text="性别:").grid(row=1, column=1, sticky=W)
        self.survey.field_sex = tk.StringVar(frame)
        self.survey.field_sex.set('1')
        sexes = ['男', '女']
        i = 1
        for sex in sexes:
            int_var = IntVar()
            temp_r = Radiobutton(frame, text=sex, value=i, variable=self.survey.field_sex, command=changed)
            temp_r.grid(row=1, column=i + 1, sticky=W)

            i = i + 1

    def init_q2(self):
        frame = self.create_new_frame()
        self.survey.field_age = self.init_input_field(frame, text='你的家在:',
                                                      val='', row=1, column=0, with_star=False, width=20)
        self.survey.field_age = self.init_input_field(frame, text='你的预期薪资:',
                                                      val='', row=1, column=3, with_star=False, width=20)
        self.survey.field_age = self.init_input_field(frame, text='家里几口人:',
                                                      val='', row=1, column=5, with_star=False, width=20)
        self.survey.field_age = self.init_input_field(frame, text='你的梦想是什么:',
                                                      val='', row=2, column=0, with_star=False, width=20)

    def init_q3(self):
        frame = self.create_new_frame(text='你生命中觉得最重要的东西是什么（最多3个）:')
        # Label(frame, text="你生命中觉得最重要的东西是什么（最多3个）:").grid(row=0, column=0, sticky=W)
        self.survey.field_sex.set(1)
        items = ['金钱', '家人', '健康', '名声', '其它.']
        self.survey.field_vips = []
        i = 0
        for item in items:
            temp = tk.StringVar(frame)
            self.survey.field_vips.append(temp)
            cb = ttk.Checkbutton(frame,
                                 text=item,
                                 variable=temp,
                                 onvalue=i,
                                 offvalue='None',
                                 command=changed)
            cb.grid(row=0, column=i, sticky=W)
            i += 1

        entry = tk.Entry(frame, width=40)
        entry.grid(row=0, column=i)

    def do_ok(self):
        pass


def gamerun():
    print('run game...')


def test(content):
    # 如果不加上==""的话，就会发现删不完。总会剩下一个数字
    if content.isdigit() or content == "":
        return True
    else:
        return False


def init_radio(root):
    frame_radio = ttk.Labelframe(root, text='', padding=20)
    frame_radio.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    books = ['男', '女']
    i = 0
    sexs_radio = []
    for book in books:
        int_var = IntVar()
        sexs_radio.append(int_var)
        Radiobutton(frame_radio, text=book, value=i, variable=int_var, command=changed).pack(side=LEFT)
        i = i + 1


def changed():
    print('value changed')


def init_checkbutton(root):
    frame_checkbutton = ttk.Labelframe(root, text='Checkbutton Test', padding=20)
    frame_checkbutton.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    books = ['C++', 'Python', 'Linux', 'Java']
    i = 0
    books_checkbox = []
    for book in books:
        strVar = StringVar()
        books_checkbox.append(strVar)
        cb = ttk.Checkbutton(frame_checkbutton,
                             text=book,
                             variable=strVar,
                             onvalue=i,
                             offvalue='None',
                             command=changed)
        cb.pack(anchor=W)
        i += 1


def init_combobox(root):
    frame_combobox = ttk.Labelframe(root, text='Combobox Test', padding=20)
    frame_combobox.pack(fill=BOTH, expand=YES, padx=10, pady=10)
    strVar = StringVar()
    # 创建Combobox组件
    cb = ttk.Combobox(frame_combobox,
                      textvariable=strVar,  # 绑定到self.strVar变量
                      postcommand=changed)  # 当用户单击下拉箭头时触发self.choose方法
    cb.pack(side=TOP)
    # 为Combobox配置多个选项
    cb['values'] = ['Python', 'Ruby', 'Kotlin', 'Swift']


def show_it():
    messagebox.showinfo(title='Alert', message="Please try again!")


def init_showinfo(root):
    frame_showinfo = ttk.LabelFrame(root, text='ShowInfoTest:')
    frame_showinfo.pack(side=TOP, fill=X)
    Button(frame_showinfo, text="Click Me", command=show_it).pack(side=LEFT, fill=Y)


def init_ui():
    SurveyUI()
    #
    # root = Tk()
    # # root.geometry('580x680+200+100')
    # root.resizable(width=True, height=True)
    # root.title("调查表")
    # root.iconbitmap('./imgs/login.ico')
    # img_game = PhotoImage(file='./imgs/game.png')
    # img_help = ImageTk.PhotoImage(file="./imgs/help.png")
    # init_menu(root, img_game, img_help)
    # init_userinfo(root)
    # # init_showinfo(root)
    # init_radio(root)
    # init_checkbutton(root)
    # init_combobox(root)
    # root.mainloop()
