import docx


class parser:
	file = None
	main_information = {'code': '', 'name' : '', 'direction' : '', 'profile' : ''}
	competitions = []

	def __init__(self, filename):
		self.file = docx.Document(filename)

	def go_parse(self):
		self.main_information["code"], self.main_information["name"] = self.file.tables[1].rows[0].cells[3].text.split(" ", 1)
		self.main_information["direction"] = self.file.tables[1].rows[2].cells[1].text
		self.main_information["profile"] = self.file.tables[1].rows[4].cells[4].text
		self.find_competitions()

		

	def print_inform(self):
		for info in self.main_information:
			print(self.main_information[info])
		print()
		print()
		for competition in self.competitions:
			for key in competition:
				if key == "indicators":
					for item in competition[key]:
						print(item[0])
						print(item[1])
				print(competition[key])
			print()

	def get_main_information(self):
		return self.main_information
		
	def get_comprtitions(self):
		return self.competitions

	def print(self):
		for row in self.file.tables[3].rows:
			for cell in row.cells:
				print(cell.text)
			print()

	def find_competitions(self):
		buf = {'index' : '', 'content' : '', 'indicators' : []}
		for row in self.file.tables[3].rows[1:]:
			buf['index'], buf['content'] = row.cells[0].text.split(" ", 1)
			buf_indicator = (row.cells[1].text, row.cells[2].text)
			buf['indicators'].append(buf_indicator)
		self.competitions.append(buf)


