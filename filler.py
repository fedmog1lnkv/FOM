from tkinter.messagebox import NO
from docx.enum.text import WD_COLOR_INDEX as colors
from turtle import width
from parser import parser
import docx


class filler:
	table_width = 5939790
	doc = None
	filename_to_parse = ""
	main_information = None

	def parsing(self, filename):
		
		local_parser = parser(filename)
		local_parser.go_parse()
		self.main_information = local_parser.get_main_information()

	def __init__(self, filename, filename_to_parse):
		self.doc = docx.Document(filename)
		self.filename_to_parse = filename_to_parse
	
	def add_competition(self, competition):
		pass

	def fill_main_information(self):
		self.parsing(self.filename_to_parse)
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
					first_run.text = self.main_information[buf[4:]]
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
								first_run.text = self.main_information[buf[4:]]
								first_run.font.highlight_color = colors.AUTO
								buf = ""
								first_run = None

	def save(self, filename):
		self.doc.save(filename)

