from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from filler import filler
from parser import parser

FQ = "1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован\n1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован"

def on_closing():
    if messagebox.askokcancel("Выход из приложения", "Хотите выйти из приложения?"):
        root.destroy()

values = ("Конспект лекций", "Глоссарий по предмету", "Тест",
		  "Устный опрос", "Доклад/презентация", "Реферат",
		  "Эссе", "Контрольная работа", "Практическое задание",
		  "Решение задач", "Лабороторная работа", "Проект",
		  "Портфолио", "Выставка", "Деловая игра",
		  "Конференция", "Олимпиада", "Онлайн - курс")


class app:

	editor = filler("docx/shablon_test.docx")
	local_parser = None
	main_information = None
	competitions = None

	def __init__(self, master):
		self.master = master
		self.master.title("ФОНД ОЦЕНОЧНЫХ СРЕДСТВ")
		self.master.geometry("800x600")
		self.page_1()

	def request_file(self):
		filename = filedialog.askopenfilename()
		self.local_parser = parser(filename)
		self.local_parser.go_parse()
		self.main_information = self.local_parser.get_main_information()
		self.competitions = self.local_parser.get_comprtitions()
		self.editor.fill_main_information(self.main_information)
		self.editor.save("docx/new_from_maket.docx")
		self.label_file = Label(self.frame1,
					text=filename,
					font=("Arial Bold", 10),
                                       fg="#000000",
                                      bg="#ffffff")
		self.label_file.place(relx=0.6, rely=0.65)

	def page_1(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame1 = Frame(self.master, bg="#E6E6E6")
		self.frame1.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.label_1 = Label(self.frame1, bg="#2A579A", fg="#eee", text="Автоматизация шаблона", font=("Arial Bold", 25), justify=CENTER)
		self.label_1.place(relwidth=1)
		self.label_2 = Label(self.frame1, bg="#E6E6E6", fg="#000",
							 text="Программа создана для\nупрощения вашей работы.\nВы сможете \nотредактировать документ \nпо шаблону, не прилагая \nмного усилий и кучу \nвремени",
							 justify=LEFT, font=("Arial Bold", 20))
		self.label_2.place(relx=0.08, rely=0.26)



		self.btn_file = Button(self.frame1, text="Укажите файл", bg="#2A579A", font=("Arial Bold", 15), fg="#eee",
							   command=self.request_file
							   )
		self.btn_file.place(relx=0.08, rely=0.7)

		self.FAQ_btn = ttk.Button(self.frame1, text="Выход", command=self.close_app)
		self.FAQ_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

		self.FAQ_btn = ttk.Button(self.frame1, text="О нас", command=self.page_3)
		self.FAQ_btn.place(relx=0.45, rely=0.91, relwidth=0.15)

		self.next_btn = ttk.Button(self.frame1, text="Далее", command=self.page_2)
		self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

	def page_2(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame2 = Frame(self.master, bg="#E6E6E6", width=300, height=300)
		self.frame2.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.label_1 = Label(self.frame2, bg="#2A579A", fg="#eee", text="РПД(На гридах потом сделаю)", font=("Arial Bold", 25), justify=CENTER)
		self.label_1.place(relwidth=1)

		self.info_name = Label(self.frame2, bg="#2A579A", text="Наименование дисциплины", fg="#eee", font=("Arial Bold", 15), justify=CENTER)
		self.info_1.place(relx=0.15, rely=0.15 )

		self.re_info_name = Label(self.frame2, bg="#F1F1F1", text="Основы программирования", fg="#000", font=("Arial Bold", 15), justify=CENTER)
		self.re_info_1.place(relx=0.55, rely=0.15)

		self.info_2 = Label(self.frame2, bg="#2A579A", text="Направление подготовки", fg="#eee", font=("Arial Bold", 15), justify=CENTER)
		self.info_2.place(relx=0.15, rely=0.25 )

		self.re_info_2 = Label(self.frame2, bg="#F1F1F1", text="Прикладная информатика", fg="#000", font=("Arial Bold", 15), justify=CENTER)
		self.re_info_2.place(relx=0.55, rely=0.25)

		self.info_3 = Label(self.frame2, bg="#2A579A", text="Направленность подготовки", fg="#eee", font=("Arial Bold", 15), justify=CENTER)
		self.info_3.place(relx=0.15, rely=0.35 )

		self.re_info_3 = Label(self.frame2, bg="#F1F1F1", text="Разработка программного обеспечения", fg="#000", font=("Arial Bold", 15), justify=CENTER)
		self.re_info_3.place(relx=0.55, rely=0.35)

		self.info_4 = Label(self.frame2, bg="#2A579A", text="Квалификация выпускника", fg="#eee", font=("Arial Bold", 15), justify=CENTER)
		self.info_4.place(relx=0.15, rely=0.45 )

		self.re_info_4 = Label(self.frame2, bg="#F1F1F1", text="Бакалавр", fg="#000", font=("Arial Bold", 15), justify=CENTER)
		self.re_info_4.place(relx=0.55, rely=0.45)



		self.FAQ_btn = ttk.Button(self.frame2, text="Назад", command=self.page_1)
		self.FAQ_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

		self.next_btn = ttk.Button(self.frame2, text="Далее", command=self.page_2_1)
		self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)


		#self.changes = ttk.Label(self.frame2, text='Внесите изменения')
		#self.changes.place(relx=0.1, rely=0.7)
		#self.submit = ttk.Button(self.frame2,
		#						 text="Подтвердить изменения",
		#						 command=self.page_4)
		#self.submit.pack()

	def page_2_1(self):
		for i in self.master.winfo_children():
			i.destroy()


# create a main Frame
		self.main_frame = Frame(self.master, bg="#E6E6E6")
		self.main_frame.pack(fill=BOTH, expand=1)

		# Create a Canvas
		self.my_canvas = Canvas(self.main_frame, bg="#E6E6E6")
		self.my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

		# Add a scrollbar to the Canvas
		self.my_scrollbar = ttk.Scrollbar(self.main_frame, orient=VERTICAL, command=self.my_canvas.yview)
		self.my_scrollbar.pack(side=RIGHT, fill=Y)

		# Configure the canvas
		self.my_canvas.configure(yscrollcommand=self.my_scrollbar.set)
		self.my_canvas.bind('<Configure>', lambda e: self.my_canvas.configure(scrollregion=self.my_canvas.bbox("all")))

		# create another Frame inside the canvas
		self.second_frame = Frame(self.my_canvas, bg="#46344E")

		# Add that new frame to a window in the canvas
		self.my_canvas.create_window((0, 0), window=self.second_frame, anchor="nw")


		self.frame2 = Frame(self.second_frame, bg="#E6E6E6", width=300, height=300)
		self.frame2.place(relx=0, rely=0, relwidth=1, relheight=1)




		self.combo_info = Label(self.second_frame, text="Lorem ipsum").grid(row=0, column=0,
																			#ipadx=10, ipady=6,
																			padx=10, pady=10)

		#self.combo = ttk.Combobox(root)
		#self.combo.pack()



		self.FAQ_btn = ttk.Button(self.frame2, text="Назад", command=self.page_2)
		self.FAQ_btn.grid(row=2, column=0, ipadx=10, ipady=6, padx=10, pady=10)








	def page_3(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame3 = Frame(self.master, width=300, height=300)
		self.frame3.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.FAQ_lbl = ttk.Label(self.frame3,
								 text=FQ,
								 justify=LEFT,
								 font=("Arial Bold", 15))
		self.FAQ_lbl.place(relx=0.05, rely=0.1)

		self.FAQ_btn = ttk.Button(self.frame3,
								  text="О нас",
								  command=self.page_3)
		self.FAQ_btn.place(relx=0.65, rely=0.91, relwidth=0.1)

		self.page1_btn = ttk.Button(self.frame3,
									text="Вернуться в начало",
									command=self.page_1)
		self.page1_btn.place(relx=0.75, rely=0.91, relwidth=0.2)

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
root.protocol("WM_DELETE_WINDOW", on_closing) #Подтверждение при закрытии
root.wm_attributes("-topmost", 1)#Выводится над всеми окнами
app(root)
root.mainloop()
