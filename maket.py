from threading import local
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
	all_forms = (
		"Конспект лекций",
		"Глоссарий по предмету",
		"Тест",
		"Устный опрос",
		"Доклад/презентация",
		"Реферат",
		"Эссе",
		"Контрольная работа",
		"Практическое задание",
		"Решение задач",
		"Лабороторная работа",
		"Проект",
		"Портфолио",
		"Выставка",
		"Деловая игра",
		"Конференция",
		"Олимпиада",
		"Онлайн - курс",
	)

	forms = []

	editor = filler("docx/shablon_test.docx")
	local_parser = None
	main_information = None
	competences = None
	evalutions_tools = []
	themes = []
	values = {
		"code": '',
		"name": '',
		"direction": '',
		"profile": '',
		"forms" : []
	}
	right_file = False

	def __init__(self, master):
		self.master = master
		self.master.title("ФОНД ОЦЕНОЧНЫХ МАТЕРИАЛОВ")
		self.master.geometry("1200x700")
		self.master.resizable(width=False, height=False)
		self.page_1()

	def request_file(self):
		filename = filedialog.askopenfilename()
		if filename != "":
			self.right_file = True
		self.local_parser = parser(filename)
		self.main_information = self.local_parser.get_main_information()
		self.competences = self.local_parser.get_competences()

		self.label_file = Label(
			self.frame1,
			text=filename,
			font=("Arial Bold", 10),
			fg="#000000",
			bg="#ffffff",
		)
		self.label_file.place(relx=0.08, rely=0.8)

	def page_1(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame1 = Frame(self.master, bg="#E6E6E6")
		self.frame1.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.label_1 = Label(
			self.frame1,
			bg="#2A579A",
			fg="#eee",
			text="Автоматизация шаблона",
			font=("Arial Bold", 25),
			justify=CENTER,
		)
		self.label_1.place(relwidth=1)
		self.label_2 = Label(
			self.frame1,
			bg="#E6E6E6",
			fg="#000",
			text="Программа создана для\nупрощения вашей работы.\nВы сможете \nотредактировать документ \nпо шаблону, не прилагая \nмного усилий и кучу \nвремени",
			justify=LEFT,
			font=("Arial Bold", 20),
		)
		self.label_2.place(relx=0.08, rely=0.26)

		self.btn_file = Button(
			self.frame1,
			text="Укажите файл",
			bg="#2A579A",
			font=("Arial Bold", 15),
			fg="#eee",
			command=self.request_file,
		)
		self.btn_file.place(relx=0.08, rely=0.7)

		self.img_main = ImageTk.PhotoImage(Image.open("image/First_page.png"))
		self.panel_main = Label(self.frame1, image=self.img_main, bg="#E6E6E6")
		self.panel_main.place(relx=0.6, rely=0.26)

		self.quit_btn = ttk.Button(self.frame1, text="Выход", command=self.close_app)
		self.quit_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

		self.FAQ_btn = ttk.Button(self.frame1, text="О нас", command=self.page_FAQ)
		self.FAQ_btn.place(relx=0.45, rely=0.91, relwidth=0.15)

		self.next_btn = ttk.Button(self.frame1, text="Далее", command=self.page_2)
		self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

	def open_page_1(self):
		self.right_file = False
		self.local_parser.clear()
		self.page_1()

	def page_2(self):
		if not(self.right_file):
			return

		self.values['forms'].clear()
		for i in self.master.winfo_children():
			i.destroy()

		self.frame2 = Frame(self.master, bg="#E6E6E6", width=300, height=300)
		self.frame2.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.back_btn = ttk.Button(self.frame2, text="Назад", command=self.open_page_1)
		self.back_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

		self.header = Label(
			self.frame2,
			bg="#2A579A",
			fg="#eee",
			text="Данные из РПД",
			font=("Arial Bold", 25),
		)
		self.header.place(relwidth=1)

		self.info_code = Label(
			self.frame2,
			bg="#2A579A",
			text="Код дисциплины",
			fg="#eee",
			font=("Arial Bold", 15),
			width="30",
		)
		self.info_code.place(relx=0.15, rely=0.15)

		self.re_info_code = Entry(
			self.frame2, bg="#F1F1F1", fg="#000", width=30, font=("Arial Bold", 15)
		)
		self.re_info_code.place(relx=0.55, rely=0.15)
		self.re_info_code.insert(0, self.main_information["code"])

		self.info_name = Label(
			self.frame2,
			bg="#2A579A",
			text="Наименование дисциплины",
			fg="#eee",
			font=("Arial Bold", 15),
			width=30,
		)
		self.info_name.place(relx=0.15, rely=0.25)

		self.re_info_name = Entry(
			self.frame2, bg="#F1F1F1", fg="#000", width=30, font=("Arial Bold", 15)
		)
		self.re_info_name.place(relx=0.55, rely=0.25)
		self.re_info_name.insert(0, self.main_information["name"])

		self.info_direction = Label(
			self.frame2,
			bg="#2A579A",
			text="Профиль подготовки",
			fg="#eee",
			font=("Arial Bold", 15),
			width=30,
		)
		self.info_direction.place(relx=0.15, rely=0.35)

		self.re_info_direction = Entry(
			self.frame2, bg="#F1F1F1", fg="#000", width=30, font=("Arial Bold", 15)
		)
		self.re_info_direction.place(relx=0.55, rely=0.35)
		self.re_info_direction.insert(0, self.main_information["direction"])

		self.info_profile = Label(
			self.frame2,
			bg="#2A579A",
			text="Профиль подготовки",
			fg="#eee",
			font=("Arial Bold", 15),
			width=30,
		)
		self.info_profile.place(relx=0.15, rely=0.45)

		self.re_info_profile = Entry(
			self.frame2, bg="#F1F1F1", fg="#000", width=30, font=("Arial Bold", 15)
		)
		self.re_info_profile.place(relx=0.55, rely=0.45)
		self.re_info_profile.insert(0, self.main_information["profile"])

		self.description_forms = Label(
			self.frame2,
			bg="#2A579A",
			text="Формы контроля:",
			fg="#eee",
			font=("Arial Bold", 15),
			width=30,
		)
		self.description_forms.place(relx=0.15, rely=0.55)

		self.list_of_forms = tkinter.Listbox(
			self.frame2,
			width=37,
			height=9,
			bg="#f1f1f1",
			font=("Arial Bold", 12),
			selectmode=tkinter.MULTIPLE
		)
		self.list_of_forms.insert(0, *self.all_forms)
		self.list_of_forms.place(relx=0.55, rely=0.55)

		self.next_btn = ttk.Button(self.frame2, text="Далее", command=self.page_3)
		self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

	def page_3(self):

		self.values = {
			"code": self.re_info_code.get(),
			"name": self.re_info_name.get(),
			"direction": self.re_info_direction.get(),
			"profile": self.re_info_profile.get(),
		}

		indexes = self.list_of_forms.curselection()
		for i in indexes:
			self.forms.append(self.all_forms[i])
		self.values["forms"] = self.forms

		for i in self.master.winfo_children():
			i.destroy()

		self.frame3 = Frame(self.master, bg="#E6E6E6")
		self.frame3.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.header = Label(
			self.frame3,
			bg="#2A579A",
			fg="#eee",
			text="Паспорт фонда",
			font=("Arial Bold", 25),
		)
		self.header.place(relwidth=1)

		self.frame_table = Frame(self.frame3, bg="#E6E6E6")
		self.frame_table.place(relx=0.05, rely=0.1, relwidth=1, relheight=0.9)

#шапка табицы
		self.number_pp = Label(
			self.frame_table,
			bg="#2A579A",
			text="№ п/п",
			fg="#eee",
			font=("Arial Bold", 15),
			width=6,
			height=2
		)
		self.number_pp.grid(row = 0, column=0, padx=6)

		self.index_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="Индекс\nкомпетенции",
			fg="#eee",
			font=("Arial Bold", 15),
			width=11
		)
		self.index_competence.grid(row = 0, column=1, padx=6)

		self.content_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="Содержание\nкомпетенции",
			fg="#eee",
			font=("Arial Bold", 15),
			width=25
		)
		self.content_competence.grid(row=0, column=2, padx=6)

		self.index_content_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="ИДК",
			fg="#eee",
			font=("Arial Bold", 15),
			width=7,
			height=2
		)
		self.index_content_competence.grid(row=0, column=3, padx=6)

		self.learning_outcomes = Label(
			self.frame_table,
			bg="#2A579A",
			text="Планируемые\nрезультаты обучения",
			fg="#eee",
			font=("Arial Bold", 15),
			width=19
		)
		self.learning_outcomes.grid(row=0, column=4, padx=6)

		self.evaluation_tool = Label(
			self.frame_table,
			bg="#2A579A",
			text="Наименование\nоценочного средства",
			fg="#eee",
			font=("Arial Bold", 15),
			width=19
		)
		self.evaluation_tool.grid(row=0, column=5, padx=6)

#таблица

		self.i = 0
		self.third_line = True

		self.compl_number_pp = Label(
			self.frame_table,
			bg="#2A579A",
			text="",
			fg="#eee",
			font=("Arial Bold", 13),
			width=6,
			height=2
		)
		self.compl_number_pp.grid(row=1, rowspan=3, column=0, padx=3)

		self.compl_index_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="",
			fg="#eee",
			font=("Arial Bold", 13),
			width=7,
            height = 2
		)
		self.compl_index_competence.grid(row=1, rowspan=3, column=1, padx=3)

		self.compl_content_competence = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=22,
			width=30,
			wrap=WORD,
		)
		self.compl_content_competence.grid(row=1, rowspan=3, column=2, padx=3)

		self.compl1_index_content_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="",
			fg="#eee",
			font=("Arial Bold", 13),
			width=8,
			height=2
		)
		self.compl1_index_content_competence.grid(row=1, column=3, padx=3)

		self.compl1_learning_outcomes = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=7,
			width=23,
			wrap=WORD
		)
		self.compl1_learning_outcomes.grid(row=1, column=4, padx=3, pady=3)

		self.compl1_evaluation_tool = ttk.Combobox(self.frame_table, values = self.values["forms"], font=("Arial Bold", 13))
		self.compl1_evaluation_tool.grid(row=1, column=5, padx=3, ipadx = 2, ipady = 2)

		self.compl2_index_content_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="",
			fg="#eee",
			font=("Arial Bold", 13),
			width=8,
			height=2
		)
		self.compl2_index_content_competence.grid(row=2, column=3, padx=3)

		self.compl2_learning_outcomes = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=7,
			width=23,
			wrap=WORD
		)
		self.compl2_learning_outcomes.grid(row=2, column=4, padx=3, pady=3)

		self.compl2_evaluation_tool = ttk.Combobox(self.frame_table, values = self.values["forms"], font=("Arial Bold", 13))
		self.compl2_evaluation_tool.grid(row=2, column=5, padx=3,ipadx = 2, ipady = 2)

		self.compl3_index_content_competence = Label(
			self.frame_table,
			bg="#2A579A",
			text="",
			fg="#eee",
			font=("Arial Bold", 13),
			width=8,
			height=2
		)
		self.compl3_index_content_competence.grid(row=3, column=3, padx=3)

		self.compl3_learning_outcomes = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=7,
			width=23,
			wrap=WORD
		)
		self.compl3_learning_outcomes.grid(row=3, column=4, padx=3, pady=3)

		self.compl3_evaluation_tool = ttk.Combobox(self.frame_table, values = self.values["forms"], font=("Arial Bold", 13))
		self.compl3_evaluation_tool.grid(row=3, column=5, padx=3, ipadx = 2, ipady = 2)

		self.back_btn = ttk.Button(self.frame3, text="Назад", command=self.page_2)
		self.back_btn.place(relx=0.15, rely=0.91, relwidth=0.15)

		self.back_table_btn = ttk.Button(self.frame_table, text="Назад по таблице", command=self.competence_to_interface_back)
		self.back_table_btn.place(relx=0.50, rely=0.91, relwidth=0.15)

		self.next_table_btn = ttk.Button(self.frame_table, text="Далее по таблице", command=self.competence_to_interface_next)
		self.next_table_btn.place(relx=0.75, rely=0.91, relwidth=0.15)

		self.write_competence(self.i)


	def write_competence(self, i):
		if i < len(self.competences) :
			self.compl_number_pp.config(text = str(self.i + 1))
			self.compl_index_competence.config(text = self.competences[i]['index'])
			self.compl_content_competence.delete("1.0","end")
			self.compl_content_competence.insert(1.0, self.competences[i]['content'])
			if len(self.competences[i]['indicators']) == 2:
				if self.third_line:
					self.compl3_index_content_competence.grid_forget()
					self.compl3_learning_outcomes.grid_forget()
					self.compl3_evaluation_tool.grid_forget()
					self.third_line = False

				self.compl1_index_content_competence.config(text = self.competences[i]['indicators'][0][0])
				self.compl2_index_content_competence.config(text = self.competences[i]['indicators'][1][0])


				self.compl1_learning_outcomes.delete("1.0","end")
				self.compl2_learning_outcomes.delete("1.0","end")
				self.compl1_learning_outcomes.insert(1.0, self.competences[i]['indicators'][0][1])
				self.compl2_learning_outcomes.insert(1.0, self.competences[i]['indicators'][1][1])

			elif len(self.competences[i]['indicators']) == 3:
				if not(self.third_line):
					self.compl3_index_content_competence.grid(row=3, column=3, padx=3)
					self.compl3_learning_outcomes.grid(row=3, column=4, padx=3, pady=3)
					self.compl3_evaluation_tool.grid(row=3, column=5, padx=3, ipadx = 2, ipady = 2)
					self.third_line = True

				self.compl1_index_content_competence.config(text=self.competences[i]['indicators'][0][0])
				self.compl2_index_content_competence.config(text=self.competences[i]['indicators'][1][0])
				self.compl3_index_content_competence.config(text=self.competences[i]['indicators'][2][0])

				self.compl1_learning_outcomes.delete("1.0","end")
				self.compl2_learning_outcomes.delete("1.0","end")
				self.compl3_learning_outcomes.delete("1.0","end")
				self.compl1_learning_outcomes.insert(1.0, self.competences[i]['indicators'][0][1])
				self.compl2_learning_outcomes.insert(1.0, self.competences[i]['indicators'][1][1])
				self.compl3_learning_outcomes.insert(1.0, self.competences[i]['indicators'][2][1])
		if i + 1 == len(self.competences):
			self.next_table_btn.config(command=self.end_page_3, text="Далее по таблице")

	def competence_to_interface_back(self):
		if self.i > 0:
			self.evalutions_tools.pop(len(self.evalutions_tools) - 1)
			self.i -= 1
			self.write_competence(self.i)

		if self.i + 1 < len(self.competences):
			self.next_table_btn.config(command=self.competence_to_interface_next, text="Далее по таблице")

	def competence_to_interface_next(self):
		if self.i < len(self.competences):
			if len(self.competences[self.i]['indicators']) == 2:
				self.evalutions_tools.append([self.compl1_evaluation_tool.get(), self.compl2_evaluation_tool.get()])
			if len(self.competences[self.i]['indicators']) == 3:
				self.evalutions_tools.append([self.compl1_evaluation_tool.get(), self.compl2_evaluation_tool.get(),
											  self.compl3_evaluation_tool.get()])
		if self.i + 1 < len(self.competences):
			self.i += 1
			self.write_competence(self.i)
			if self.i + 1 == len(self.competences):
				self.next_table_btn.config(command=self.end_page_3, text="Далее")
		else:
			self.next_table_btn.config(command=self.end_page_3, text="Далее")



	def end_page_3(self):
		if self.i < len(self.competences):
			if len(self.competences[self.i]['indicators']) == 2:
				self.evalutions_tools.append([self.compl1_evaluation_tool.get(), self.compl2_evaluation_tool.get()])
			if len(self.competences[self.i]['indicators']) == 3:
				self.evalutions_tools.append([self.compl1_evaluation_tool.get(), self.compl2_evaluation_tool.get(),
											  self.compl3_evaluation_tool.get()])
		self.page_4()

	def page_4(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.idks = self.local_parser.get_idks()
		self.idks_to_table = [self.idks[i][0] for i in range(len(self.idks))]
		self.themes = self.local_parser.get_themes()
		self.frame4 = Frame(self.master, bg="#E6E6E6")
		self.frame4.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.header = Label(
			self.frame4,
			bg="#2A579A",
			fg="#eee",
			text="Показатели и критерии оценивания компетенций",
			font=("Arial Bold", 25),
		)
		self.header.place(relwidth=1)

		self.frame_table = Frame(self.frame4, bg="#E6E6E6")
		self.frame_table.place(relx=0.05, rely=0.1, relwidth=1, relheight=0.9)

# шапка табицы

		self.topic_discipline = Label(
			self.frame_table,
			bg="#2A579A",
			text="Тема или раздел\nдисциплины",
			fg="#eee",
			font=("Arial Bold", 15),
			height=4,
			width=15
		)
		self.topic_discipline.grid(row=0, rowspan=2, column=0, padx=6)

		self.competence_indicator = Label(
			self.frame_table,
			bg="#2A579A",
			text="Код индикатора\nкомпетенции",
			fg="#eee",
			font=("Arial Bold", 15),
			height=4,
			width=20
		)
		self.competence_indicator.grid(row=0, rowspan=2, column=1, padx=6)

		self.planned_result = Label(
			self.frame_table,
			bg="#2A579A",
			text="Планируемый\nрезультат",
			fg="#eee",
			font=("Arial Bold", 15),
			height=4,
			width=13
		)
		self.planned_result.grid(row=0, rowspan=2, column=2, padx=6)

		self.indicator = Label(
			self.frame_table,
			bg="#2A579A",
			text="Показатель",
			fg="#eee",
			font=("Arial Bold", 15),
			height=4,
			width=11
		)
		self.indicator.grid(row=0, rowspan=2, column=3, padx=6)

		self.evaluation_criteria = Label(
			self.frame_table,
			bg="#2A579A",
			text="Критерий\nоценивания",
			fg="#eee",
			font=("Arial Bold", 15),
			height=4,
			width=11
		)
		self.evaluation_criteria.grid(row=0, rowspan=2, column=4, padx=6)

		self.name_os = Label(
			self.frame_table,
			bg="#2A579A",
			text="Наименование ОС",
			fg="#eee",
			font=("Arial Bold", 15),
			height=2,
			width=18
		)
		self.name_os.grid(row=0, column=5, columnspan=2, padx=6)

		self.tk = Label(
			self.frame_table,
			bg="#2A579A",
			text="ТК",
			fg="#eee",
			font=("Arial Bold", 15),
			height=2,
			width=8
		)
		self.tk.grid(row=1, column=5, padx=0)

		self.pa = Label(
			self.frame_table,
			bg="#2A579A",
			text="ПА",
			fg="#eee",
			font=("Arial Bold", 15),
			height=2,
			width=8
		)
		self.pa.grid(row=1, column=6, padx=0)

#таблица
		self.i = 0
		self.compl_topic_discipline = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=10,
			width=18,
			wrap=WORD
		)
		self.compl_topic_discipline.grid(row=2, column=0, padx=3)

		self.compl_competence_indicators = tkinter.Listbox(
			self.frame_table,
			width=25,
			height=10,
			bg="#f1f1f1",
			font=("Arial Bold", 12),
			selectmode=tkinter.MULTIPLE,
		)
		self.compl_competence_indicators.insert(0, *self.idks_to_table)
		self.compl_competence_indicators.grid(row=2, column=1, padx=3)

		self.compl_planned_result = Label(
			self.frame_table,
			bg="#2A579A",
			text="hohoho,\nNO.",
			fg="#eee",
			font=("Arial Bold", 13),
			height=10,
			width=16
		)
		self.compl_planned_result.grid(row=2, column=2, padx=3)

		self.compl_indicator = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=10,
			width=13,
			wrap=WORD
		)
		self.compl_indicator.grid(row=2, column=3, padx=3)

		self.compl_evaluation_criteria = Text(self.frame_table,
			font=("Arial Bold", 13),
			height=10,
			width=13,
			wrap=WORD
		)
		self.compl_evaluation_criteria.grid(row=2, column=4, padx=3)

		self.compl_tk = ttk.Combobox(self.frame_table, values=self.values["forms"], font=("Arial Bold", 13), height=10, width=7)
		self.compl_tk.grid(row=2, column=5, padx=3)

		self.compl_pa = ttk.Combobox(self.frame_table, values=self.values["forms"], font=("Arial Bold", 13), height=10, width=7)
		self.compl_pa.grid(row=2, column=6, padx=3)
		self.write_themes(self.i)

		self.back_btn = ttk.Button(self.frame4, text="Назад", command=self.page_2)
		self.back_btn.place(relx=0.15, rely=0.91, relwidth=0.15)

		self.back_table_themes_btn = ttk.Button(self.frame_table, text="Назад по таблице", command=self.themes_table_back)
		self.back_table_themes_btn.place(relx=0.50, rely=0.91, relwidth=0.15)

		self.next_table_themes_btn = ttk.Button(self.frame_table, text="Далее по таблице", command=self.themes_table_next)
		self.next_table_themes_btn.place(relx=0.75, rely=0.91, relwidth=0.15)

	def write_themes(self, i):
		self.compl_topic_discipline.delete("1.0","end")
		self.compl_topic_discipline.insert(1.0, self.themes[i]['name'])

	def themes_table_next(self):
		if self.i < len(self.themes):
			self.themes[self.i]['indicators'] = [self.idks[i][0] for i in self.compl_competence_indicators.curselection()]
			self.themes[self.i]['planned_result'] = [self.idks[i][1] for i in self.compl_competence_indicators.curselection()]
			self.themes[self.i]['indicator'] = [self.compl_indicator.get("1.0","end")]
			self.themes[self.i]['evaluation_criteria'] = [self.compl_evaluation_criteria.get("1.0","end")]
			self.themes[self.i]['name_os'] = [self.compl_tk.get()]
			self.themes[self.i]['tk'] = [self.compl_tk.get()]

		if self.i + 1 < len(self.themes):
			self.i += 1
			self.write_themes(self.i)
			if self.i + 1 == len(self.themes):
				self.next_table_themes_btn.config(command = self.end_page_4, text = "Далее")
		else:
			self.next_table_themes_btn.config(command = self.end_page_4, text = "Далее")

	def themes_table_back(self):
		if self.i > 0:
			for key in self.themes[self.i]:
				if key != 'index' and key !='name':
					self.themes[self.i][key].clear()
			self.i -= 1
			self.write_themes(self.i)

		if self.i + 1 < len(self.competences):
			self.next_table_themes_btn.config(command=self.themes_table_next, text="Далее по таблице")

		if self.i + 1 < len(self.competences):
			self.next_table_themes_btn.config(command=self.themes_table_next, text="Далее по таблице")

	def end_page_4(self):
		if self.i < len(self.themes):
			self.themes[self.i]['indicators'] = [self.idks[i][0] for i in self.compl_competence_indicators.curselection()]
			self.themes[self.i]['planned_result'] = [self.idks[i][1] for i in self.compl_competence_indicators.curselection()]
			self.themes[self.i]['indicator'] = [self.compl_indicator.get("1.0","end")]
			self.themes[self.i]['evaluation_criteria'] = [self.compl_evaluation_criteria.get("1.0","end")]
			self.themes[self.i]['name_os'] = [self.compl_tk.get()]
			self.themes[self.i]['tk'] = [self.compl_tk.get()]
		self.page_5()

	def page_5(self):
		for i in self.master.winfo_children():
			i.destroy()
		
		for i in range(len(self.competences)):
			for j in range(len(self.competences[i]['indicators'])):
				self.competences[i]['indicators'][j] = (self.competences[i]['indicators'][j][0], self.competences[i]['indicators'][j][1], self.evalutions_tools[i][j])

		self.frame4 = Frame(self.master, width=300, height=300)
		self.frame4.place(relx=0, rely=0, relwidth=1, relheight=1)
		#self.message = ttk.Label(
		#	self.frame4, text="Сохранение файла", font=("Arial Bold", 25)
		#)

		
		self.header_save = Label(
			self.frame4,
			bg="#2A579A",
			fg="#eee",
			text="Сохранение файла",
			font=("Arial Bold", 25),
		)
		self.header_save.place(relwidth=1)
                
                #self.textsave = Label(self.frame4, text = "Введите\nназвание файла",  font=("Arial Bold", 15), justify=LEFT)
                #self.textsave.place(relx= 0.05, rely=0.5)

		
		#self.message.pack(expand=1)
		self.new_file_name = Entry(
			self.frame4, font=("Arial Bold", 25)
		)

		
		self.label_save = Label(
			self.frame4,
			bg="#F0F0F0",
			fg="#000",
			text="Введите\nназвание файла",
			justify=LEFT,
			font=("Arial Bold", 20),
		)
		self.label_save.place(relx=0.05, rely=0.37)

                
		self.img_save = ImageTk.PhotoImage(Image.open("image/save.png"))
		self.panel_save = Label(self.frame4, image=self.img_save, bg="#F0F0F0")
		self.panel_save.place(relx=0.6, rely=0.22)
		
		self.new_file_name.place(relx = 0.05, rely = 0.5)
		self.page1_btn = ttk.Button(
			self.frame4, text="Сохранить", command=self.save_file
		)
		self.page1_btn.place(relx=0.29, rely=0.56)
		self.page1_btn = ttk.Button(
			self.frame4, text="Вернуться в начало", command=self.open_page_1
		)
		self.page1_btn.place(relx=0.65, rely=0.91, relwidth=0.15)
		#self.page1_btn.pack()
		self.quit_btn = ttk.Button(self.frame4, text="Выход", command=self.close_app)
		#self.quit_btn.pack()
		self.quit_btn.place(relx=0.25, rely=0.91, relwidth=0.15)
		
	def save_file(self):
		if self.new_file_name.get() != "":
			if not(".docx" in self.new_file_name.get()) and "." in self.new_file_name.get():
				self.label_save.config(text = "Укажите расширение docx")
			else:
				file_name = filedialog.askdirectory() + "/" + self.new_file_name.get()
				if not(".docx" in self.new_file_name.get()) and not("." in self.new_file_name.get()):
					file_name += ".docx"
				self.save_all()
				self.editor.save(file_name)
				self.page_6()
		else:
			self.label_save.config(text = "Вы не указали имя файла")

	def save_all(self):
		self.editor.fill_main_information(self.values)
		self.editor.fill_developers(self.local_parser.get_developers())
		self.editor.fill_competences(self.competences)
		self.editor.fill_themes(self.themes)

	def page_6(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame4 = Frame(self.master, width=300, height=300)
		self.frame4.place(relx=0, rely=0, relwidth=1, relheight=1)
		self.new_file = ttk.Label(
			self.frame4, text="Файл создан", font=("Arial Bold", 25)
		)
		self.new_file.pack(expand=1)
		self.page1_btn = ttk.Button(
			self.frame4, text="Вернуться в начало", command=self.open_page_1
		)
		self.page1_btn.pack()
		self.quit_btn = ttk.Button(self.frame4, text="Выход", command=self.close_app)
		#self.quit_btn.pack()



                #self.quit_btn = ttk.Button(self.frame4, text="Выход", command=self.close_app)
		#self.quit_btn.place(relx=0.25, rely=0.91, relwidth=0.15)


		#self.next_btn = ttk.Button(self.frame4, text="Далее", command=self.page_2)
		#self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)
		

	def page_FAQ(self):
		for i in self.master.winfo_children():
			i.destroy()

		self.frame3 = Frame(self.master, width=300, height=300, bg="#E6E6E6")
		self.frame3.place(relx=0, rely=0, relwidth=1, relheight=1)

		self.label_1 = Label(
			self.frame3,
			bg="#2A579A",
			fg="#eee",
			text="О нас",
			font=("Arial Bold", 25),
			justify=CENTER,
		)
		self.label_1.place(relwidth=1)

		self.label_2 = Label(
			self.frame3,
			bg="#E6E6E6",
			fg="#000",
			text="Программа создана для\nупрощения вашей работы.\nВы сможете \nотредактировать документ \nпо шаблону, не прилагая \nмного усилий и кучу \nвремени",
			justify=LEFT,
			font=("Arial Bold", 20),
		)
		self.label_2.place(relx=0.08, rely=0.26)
		#########################################################################
		self.img_1 = ImageTk.PhotoImage(Image.open("image/Website.png"))
		self.panel_1 = Label(self.frame3, image=self.img_1, bg="#E6E6E6")
		self.panel_1.place(relx=0.6, rely=0.26)

		self.image_text_1 = Label(
			self.frame3,
			text="Создание\nвеб - сайтов",
			bg="#E6E6E6",
			fg="#000",
			font=("Arial Bold", 15),
			justify=CENTER,
		)
		self.image_text_1.place(relx=0.58, rely=0.41)
		####################################################################
		self.img_2 = ImageTk.PhotoImage(Image.open("image/Website.png"))
		self.panel_2 = Label(self.frame3, image=self.img_2, bg="#E6E6E6")
		self.panel_2.place(relx=0.8, rely=0.26)

		self.image_text_2 = Label(
			self.frame3,
			text="Создание\nвеб - сайтов",
			bg="#E6E6E6",
			fg="#000",
			font=("Arial Bold", 15),
			justify=CENTER,
		)
		self.image_text_2.place(relx=0.78, rely=0.41)
		################################################################
		#########################################################################
		self.img_3 = ImageTk.PhotoImage(Image.open("image/Website.png"))
		self.panel_3 = Label(self.frame3, image=self.img_3, bg="#E6E6E6")
		self.panel_3.place(relx=0.6, rely=0.56)

		self.image_text_3 = Label(
			self.frame3,
			text="Создание\nвеб - сайтов",
			bg="#E6E6E6",
			fg="#000",
			font=("Arial Bold", 15),
			justify=CENTER,
		)
		self.image_text_3.place(relx=0.58, rely=0.71)
		####################################################################
		self.img_4 = ImageTk.PhotoImage(Image.open("image/Website.png"))
		self.panel_4 = Label(self.frame3, image=self.img_4, bg="#E6E6E6")
		self.panel_4.place(relx=0.8, rely=0.56)

		self.image_text_4 = Label(
			self.frame3,
			text="Создание\nвеб - сайтов",
			bg="#E6E6E6",
			fg="#000",
			font=("Arial Bold", 15),
			justify=CENTER,
		)
		self.image_text_4.place(relx=0.78, rely=0.71)
		################################################################
		# self.FAQ_lbl = ttk.Label(self.frame3,
		# 						 text=FQ,
		# 						 justify=LEFT,
		# 						 font=("Arial Bold", 15))
		# self.FAQ_lbl.place(relx=0.05, rely=0.1)

		# self.FAQ_btn = ttk.Button(self.frame3,
		# 						  text="О нас",
		# 						  command=self.page_3)
		# self.FAQ_btn.place(relx=0.65, rely=0.91, relwidth=0.1)

		# self.page1_btn = ttk.Button(self.frame3,
		# 							text="Вернуться в начало",
		# 							command=self.page_1)
		# self.page1_btn.place(relx=0.75, rely=0.91, relwidth=0.2)

		self.FAQ_btn = ttk.Button(self.frame3, text="Выход", command=self.close_app)
		self.FAQ_btn.place(relx=0.25, rely=0.91, relwidth=0.15)

		self.next_btn = ttk.Button(
			self.frame3, text="Вернуться в начало", command=self.page_1
		)
		self.next_btn.place(relx=0.65, rely=0.91, relwidth=0.15)

	def close_app(self):
		root.destroy()



root = Tk()
root.protocol("WM_DELETE_WINDOW", on_closing)  # Подтверждение при закрытии
root.wm_attributes("-topmost", 1)  # Выводится над всеми окнами
app(root)
root.mainloop()
