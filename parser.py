import docx


class parser:
	file = None
	information = {'code': '', 'name' : '', 'direction' : '', 'profile' : ''}

	def __init__(self, filename):
		self.file = docx.Document(filename)

	def go_parse(self):
		self.information["code"], self.information["name"] = self.file.tables[1].rows[0].cells[3].text.split(" ", 1)
		self.information["direction"] = self.file.tables[1].rows[2].cells[1].text
		self.information["profile"] = self.file.tables[1].rows[4].cells[4].text
		self.print_inform()

	def print_inform(self):
		for info in self.information:
			print(self.information[info])

	def get_information(self):
		return self.information


a = parser("rpd.docx")
a.go_parse()