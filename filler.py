from tkinter.messagebox import NO
from docx.enum.text import WD_COLOR_INDEX as colors
from turtle import width
from parser import parser
import docx


class filler:
	table_width = 5939790
	doc = None

	def __init__(self, filename):
		self.doc = docx.Document(filename)
	
	def add_competition(self, competition):
		pass

	def fill_main_information(self, main_information):
		buf = ""
		first_run = None
		for paragraph in self.doc.paragraphs:
			for run in paragraph.runs:
				if run.font.highlight_color == colors.BRIGHT_GREEN:
					buf += run.text
					if (not(first_run)):
						first_run = run
					run.text = ""
				if run.font.highlight_color != colors.BRIGHT_GREEN and buf != "":
					first_run.text = main_information[buf[4:]]
					first_run.font.highlight_color = colors.AUTO
					buf = ""
					first_run = None

		for table in self.doc.tables:
			for row in table.rows:
				for cell in row.cells:
					for paragraph in cell.paragraphs:
						for run in paragraph.runs:
							if run.font.highlight_color == colors.BRIGHT_GREEN:
								buf += run.text
								if (not(first_run)):
									first_run = run
								run.text = ""
							if run.font.highlight_color != colors.BRIGHT_GREEN and buf != "":
								first_run.text = main_information[buf[4:]]
								first_run.font.highlight_color = colors.AUTO
								buf = ""
								first_run = None

	def save(self, filename):
		self.doc.save(filename)

