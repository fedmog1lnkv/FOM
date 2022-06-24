from msilib.schema import tables
from tkinter.messagebox import NO
from docx.enum.text import WD_COLOR_INDEX as colors
from turtle import width

from pip import main
from parser import parser
import docx


class filler:
	table_width = 5939790
	doc = None

	def __init__(self, filename):
		self.doc = docx.Document(filename)
	
	def fill_main_information(self, main_information):
		buf = ""
		first_run = None
		blue_run = None
		for paragraph in self.doc.paragraphs:
			for run in paragraph.runs:
				if run.font.highlight_color == colors.BLUE:
					if not(blue_run):
						blue_run = run
						run.font.highlight_color = colors.AUTO
					run.text = ""
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

		for form in main_information['forms'][:-1]:
			blue_run.text += form.lower() + ", "
		blue_run.text += main_information['forms'][-1].lower()

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

	def fill_competitions(self, competitions):
		table = self.doc.tables[4]
		for j in range(len(competitions)):
			self.make_new_competition_section(table, len(competitions[j]['indicators']))
			n = 0
			for i in range(len(table.rows) - len(competitions[j]['indicators']), len(table.rows)):
				table.rows[i].cells[0].text = str(j + 1)
				table.rows[i].cells[1].text = competitions[j]["index"]
				table.rows[i].cells[2].text = competitions[j]["content"]
				table.rows[i].cells[3].text = competitions[j]["indicators"][n][0]
				table.rows[i].cells[4].text = competitions[j]["indicators"][n][1]
				n += 1

	def make_new_competition_section(self, table, size):
		for i in range(size):
			table.add_row()
			if i > 0:
				table.rows[-1].cells[0].merge(table.rows[-2].cells[0])
				table.rows[-1].cells[1].merge(table.rows[-2].cells[1])
				table.rows[-1].cells[2].merge(table.rows[-2].cells[2])

	def save(self, filename):
		self.doc.save(filename)
