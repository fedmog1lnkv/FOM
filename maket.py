from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from filler import filler
from parser import parser

FQ = "1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован\n1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован"


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
		self.editor.fill_competitions(self.competitions)
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

		self.frame1 = Frame(self.master)
		self.frame1.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.label_1 = Label(self.frame1,
							 text="Автоматизация шаблона",
							 font=("Arial Bold", 25))
		self.label_1.place(relx=0.28)
		self.label_2 = Label(
			self.frame1,
			text=
			"Загрузка пустого шаблона ФОМ. \nФормирование готового ФОМ\nпо заполненным разделам.",
			justify=LEFT,
			font=("Arial Bold", 20))
		self.label_2.place(relx=0.05, rely=0.1)

		self.label_5 = Label(
			self.frame1,
			text=
			"1. Выберите файл\n2. Перейдите во вкладку 'Редактирование'\n3. Сделайте замену и подтвердите\n4. Всё готово, документ отредактирован",
			justify=LEFT,
			font=("Arial Bold", 15))
		self.label_5.place(relx=0.05, rely=0.5)

		self.btn_file = Button(self.frame1,
							   text="Выберите файл",
							   bg="#7D6F86",
							   font=("Arial Bold", 15),
							   fg="#eee",
							   command=self.request_file)
		self.btn_file.place(relx=0.7, rely=0.55)


		self.next_btn = ttk.Button(self.frame1,
								   text="Далее",
								   command=self.page_2)
		self.next_btn.place(relx=0.85, rely=0.91, relwidth=0.1)

		self.FAQ_btn = ttk.Button(self.frame1,
								  text="О нас",
								  command=self.page_3)
		self.FAQ_btn.place(relx=0.75, rely=0.91, relwidth=0.1)

	def page_2(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame2 = Frame(self.master, width=300, height=300)
		self.frame2.place(relx=0, rely=0, relwidth=1, relheight=1)
		self.changes = ttk.Label(self.frame2, text='Внесите изменения')
		self.changes.pack()
		self.submit = ttk.Button(self.frame2,
								 text="Подтвердить изменения",
								 command=self.page_4)
		self.submit.pack()

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
app(root)
root.mainloop()
