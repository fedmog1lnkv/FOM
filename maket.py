from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import tkinter
from PIL import ImageTk, Image

from filler import filler
from parser import parser

FQ = "1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован\n1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован"


def on_closing():
    if messagebox.askokcancel("Выход из приложения", "Хотите выйти из приложения?"):
        root.destroy()


class app:
    all_forms = ("Конспект лекций", "Глоссарий по предмету", "Тест",
                 "Устный опрос", "Доклад/презентация", "Реферат",
                 "Эссе", "Контрольная работа", "Практическое задание",
                 "Решение задач", "Лабороторная работа", "Проект",
                 "Портфолио", "Выставка", "Деловая игра",
                 "Конференция", "Олимпиада", "Онлайн - курс")

    forms = []

    editor = filler("docx/shablon_test.docx")
    local_parser = None
    main_information = None
    competitions = None

    def __init__(self, master):
        self.master = master
        self.master.title("ФОНД ОЦЕНОЧНЫХ МАТЕРИАЛОВ")
        self.master.geometry("800x600")
        self.page_1()

    def request_file(self):
        filename = filedialog.askopenfilename()
        self.local_parser = parser(filename)
        self.main_information = self.local_parser.get_main_information()
        self.competitions = self.local_parser.get_comprtitions()

        self.label_file = Label(self.frame1,
                                text=filename,
                                font=("Arial Bold", 10),
                                fg="#000000",
                                bg="#ffffff")
        self.label_file.place(relx=0.08, rely=0.8)

    def page_1(self):
        for i in self.master.winfo_children():
            i.destroy()

        self.frame1 = Frame(self.master, bg="#E6E6E6")
        self.frame1.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.label_1 = Label(self.frame1, bg="#2A579A", fg="#eee", text="Автоматизация шаблона",
                             font=("Arial Bold", 25), justify=CENTER)
        self.label_1.place(relwidth=1)
        self.label_2 = Label(self.frame1, bg="#E6E6E6", fg="#000",
                             text="Программа создана для\nупрощения вашей работы.\nВы сможете \nотредактировать документ \nпо шаблону, не прилагая \nмного усилий и кучу \nвремени",
                             justify=LEFT, font=("Arial Bold", 20))
        self.label_2.place(relx=0.08, rely=0.26)

        self.btn_file = Button(self.frame1, text="Укажите файл", bg="#2A579A", font=("Arial Bold", 15), fg="#eee",
                               command=self.request_file
                               )
        self.btn_file.place(relx=0.08, rely=0.7)

        self.img_main = ImageTk.PhotoImage(Image.open("image/First_page.png"))
        self.panel_main = Label(self.frame1, image=self.img_main, bg="#E6E6E6")
        self.panel_main.place(relx=0.6, rely=0.26)

        self.quit_btn = ttk.Button(self.frame1, text="Выход", command=self.close_app)
        self.quit_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

        self.FAQ_btn = ttk.Button(self.frame1, text="О нас", command=self.page_3)
        self.FAQ_btn.place(relx=0.45, rely=0.91, relwidth=0.15)

        self.next_btn = ttk.Button(self.frame1, text="Далее", command=self.page_2)
        self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

    def page_2(self):
        for i in self.master.winfo_children():
            i.destroy()

        self.frame2 = Frame(self.master, bg="#E6E6E6", width=300, height=300)
        self.frame2.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.back_btn = ttk.Button(self.frame2, text="Назад", command=self.page_1)
        self.back_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

        self.label_1 = Label(self.frame2, bg="#2A579A", fg="#eee", text="Данные из РПД",
                             font=("Arial Bold", 25))
        self.label_1.place(relwidth=1)

        self.info_code = Label(self.frame2, bg="#2A579A", text="Код дисциплины", fg="#eee", font=("Arial Bold", 15),
                               width="25")
        self.info_code.place(relx=0.15, rely=0.15)

        self.re_info_code = Entry(self.frame2, bg="#F1F1F1", fg="#000", width=24, font=("Arial Bold", 15))
        self.re_info_code.place(relx=0.55, rely=0.15)
        self.re_info_code.insert(0, self.main_information["code"])

        self.info_name = Label(self.frame2, bg="#2A579A", text="Наименование дисциплины", fg="#eee",
                               font=("Arial Bold", 15), width=25)
        self.info_name.place(relx=0.15, rely=0.25)

        self.re_info_name = Entry(self.frame2, bg="#F1F1F1", fg="#000", width=24, font=("Arial Bold", 15))
        self.re_info_name.place(relx=0.55, rely=0.25)
        self.re_info_name.insert(0, self.main_information["name"])

        self.info_direction = Label(self.frame2, bg="#2A579A", text="Профиль подготовки", fg="#eee",
                                    font=("Arial Bold", 15), width=25)
        self.info_direction.place(relx=0.15, rely=0.35)

        self.re_info_direction = Entry(self.frame2, bg="#F1F1F1", fg="#000", width=24, font=("Arial Bold", 15))
        self.re_info_direction.place(relx=0.55, rely=0.35)
        self.re_info_direction.insert(0, self.main_information["direction"])

        self.info_profile = Label(self.frame2, bg="#2A579A", text="Профиль подготовки", fg="#eee",
                                  font=("Arial Bold", 15), width=25)
        self.info_profile.place(relx=0.15, rely=0.45)

        self.re_info_profile = Entry(self.frame2, bg="#F1F1F1", fg="#000", width=24, font=("Arial Bold", 15))
        self.re_info_profile.place(relx=0.55, rely=0.45)
        self.re_info_profile.insert(0, self.main_information["profile"])

        self.description_forms = Label(self.frame2, bg="#2A579A", text="Формы контроля:", fg="#eee",
                                  font=("Arial Bold", 15), width=25)
        self.description_forms.place(relx=0.15, rely=0.55)

        self.list_of_forms = tkinter.Listbox(self.frame2, width=29, height=9, bg="#f1f1f1", font=("Arial Bold", 12), selectmode=tkinter.MULTIPLE)
        self.list_of_forms.insert(0, *self.all_forms)
        self.list_of_forms.place(relx=0.55, rely=0.55)

        self.next_btn = ttk.Button(self.frame2, text="Далее", command=self.page_3)
        self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

    def page_3(self):


        values = {'code' : self.re_info_code.get(), 'name' : self.re_info_name.get(), 'direction' : self.re_info_direction.get(), 'profile' : self.re_info_profile.get()}

        indexes = self.list_of_forms.curselection()
        for i in indexes:
            self.forms.append(self.all_forms[i])
        values['forms'] = self.forms

        self.editor.fill_main_information(values)
        self.editor.fill_developers(self.local_parser.get_developers())
        self.editor.save("docx/new_from_maket.docx")

        for i in self.master.winfo_children():
            i.destroy()



    def page_4(self):
        for i in self.master.winfo_children():
            i.destroy()

        self.frame3 = Frame(self.master, width=300, height=300, bg="#E6E6E6")
        self.frame3.place(relx=0, rely=0, relwidth=1, relheight=1)

        self.label_1 = Label(self.frame3, bg="#2A579A", fg="#eee", text="О нас",
                             font=("Arial Bold", 25), justify=CENTER)
        self.label_1.place(relwidth=1)

        self.label_2 = Label(self.frame3, bg="#E6E6E6", fg="#000",
                             text="Программа создана для\nупрощения вашей работы.\nВы сможете \nотредактировать документ \nпо шаблону, не прилагая \nмного усилий и кучу \nвремени",
                             justify=LEFT, font=("Arial Bold", 20))
        self.label_2.place(relx=0.08, rely=0.26)
        #########################################################################
        self.img_1 = ImageTk.PhotoImage(Image.open("image/Website.png"))
        self.panel_1 = Label(self.frame3, image=self.img_1, bg="#E6E6E6")
        self.panel_1.place(relx=0.6, rely=0.26)

        self.image_text_1 = Label(self.frame3, text="Создание\nвеб - сайтов", bg="#E6E6E6", fg="#000",
                                  font=("Arial Bold", 15), justify=CENTER)
        self.image_text_1.place(relx=0.58, rely=0.41)
        ####################################################################
        self.img_2 = ImageTk.PhotoImage(Image.open("image/Website.png"))
        self.panel_2 = Label(self.frame3, image=self.img_2, bg="#E6E6E6")
        self.panel_2.place(relx=0.8, rely=0.26)

        self.image_text_2 = Label(self.frame3, text="Создание\nвеб - сайтов", bg="#E6E6E6", fg="#000",
                                  font=("Arial Bold", 15), justify=CENTER)
        self.image_text_2.place(relx=0.78, rely=0.41)
        ################################################################
        #########################################################################
        self.img_3 = ImageTk.PhotoImage(Image.open("image/Website.png"))
        self.panel_3 = Label(self.frame3, image=self.img_3, bg="#E6E6E6")
        self.panel_3.place(relx=0.6, rely=0.56)

        self.image_text_3 = Label(self.frame3, text="Создание\nвеб - сайтов", bg="#E6E6E6", fg="#000",
                                  font=("Arial Bold", 15), justify=CENTER)
        self.image_text_3.place(relx=0.58, rely=0.71)
        ####################################################################
        self.img_4 = ImageTk.PhotoImage(Image.open("image/Website.png"))
        self.panel_4 = Label(self.frame3, image=self.img_4, bg="#E6E6E6")
        self.panel_4.place(relx=0.8, rely=0.56)

        self.image_text_4 = Label(self.frame3, text="Создание\nвеб - сайтов", bg="#E6E6E6", fg="#000",
                                  font=("Arial Bold", 15), justify=CENTER)
        self.image_text_4.place(relx=0.78, rely=0.71)
        ################################################################
        # self.FAQ_lbl = ttk.Label(self.frame3,
        #						 text=FQ,
        #						 justify=LEFT,
        #						 font=("Arial Bold", 15))
        # self.FAQ_lbl.place(relx=0.05, rely=0.1)

        # self.FAQ_btn = ttk.Button(self.frame3,
        #						  text="О нас",
        #						  command=self.page_3)
        # self.FAQ_btn.place(relx=0.65, rely=0.91, relwidth=0.1)

        # self.page1_btn = ttk.Button(self.frame3,
        #							text="Вернуться в начало",
        #							command=self.page_1)
        # self.page1_btn.place(relx=0.75, rely=0.91, relwidth=0.2)

        self.FAQ_btn = ttk.Button(self.frame3, text="Выход", command=self.close_app)
        self.FAQ_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

        self.next_btn = ttk.Button(self.frame3, text="Вернуться в начало", command=self.page_1)
        self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

    def page_4(self):
        for i in self.master.winfo_children():
            i.destroy()

        self.frame4 = Frame(self.master, width=300, height=300)
        self.frame4.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.new_file = ttk.Label(self.frame4,
                                  text='Файл создан',
                                  font=("Arial Bold", 25))
        self.new_file.pack(expand=1)
        self.page1_btn = ttk.Button(self.frame4,
                                    text="Вернуться в начало",
                                    command=self.page_1)
        self.page1_btn.pack()
        self.quit_btn = ttk.Button(self.frame4,
                                   text="Выход",
                                   command=self.close_app)
        self.quit_btn.pack()

    def close_app(n):
        root.destroy()

    def resize(self, event):
        region = self.canvas.bbox(tk.ALL)
        self.canvas.configure(scrollregion=region)


root = Tk()
root.protocol("WM_DELETE_WINDOW", on_closing)  # Подтверждение при закрытии
root.wm_attributes("-topmost", 1)  # Выводится над всеми окнами
app(root)
root.mainloop()
