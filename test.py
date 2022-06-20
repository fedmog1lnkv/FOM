from inspect import CO_ASYNC_GENERATOR
import docx
from docx.enum.text import WD_COLOR_INDEX as colors

class auto_filler:

	def __init__(self, filename):
		self.doc = docx.Document(filename)
	
	def change(self, run, str):
		run.text = str
	
	def fill(self):
		buf = ""
		first_run = None
		for paragraph in self.doc.paragraphs:
			for run in paragraph.runs:
				if run.font.highlight_color == colors.YELLOW:
					buf += run.text
					if not(first_run):
						first_run = run
					run.text = ""
				if run.font.highlight_color != colors.YELLOW and buf != "":
					self.change(first_run, input(buf + ": "))
					buf = ""
					first_run.font.highlight_color = colors.GREEN
					first_run = None

	def save(self, filename):
		self.doc.save(filename)


file = auto_filler("shablon.docx")

file.fill()
file.save("new_file.docx")