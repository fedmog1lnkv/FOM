from tkinter import *
from tkinter import ttk

class app:
    def __init__(self, master):
        self.master = master
        self.master.title("ФОНД ОЦЕНОЧНЫХ СРЕДСТВ")
        self.master.geometry("700x500")
        self.page_1()
    
    def page_1(self):
        for i in self.master.winfo_children():
            i.destroy()
            
        self.frame1 = Frame(self.master, width=300, height=300)
        self.frame1.pack()
        self.header = ttk.Label(self.frame1, font="Arial, 22", text='Автоматическая генерация документов')
        self.header.pack()
        self.explanation = ttk.Label(self.frame1, text='Приложение для ускорения разработки и исправления фонда оценочных материалов для кафедры естественно-научных дисциплин.')
        self.explanation.pack()
        self.register_btn = ttk.Button(self.frame1, text="Далее", command=self.page_2)
        self.register_btn.pack()
    
    def page_2(self):
        for i in self.master.winfo_children():
            i.destroy()

        self.frame2 = Frame(self.master, width=300, height=300)
        self.frame2.pack()
        self.changes = ttk.Label(self.frame2, text='Внесите изменения')
        self.changes.pack()
        self.submit = ttk.Button(self.frame2, text="Подтвердить изменения", command=self.page_3)
        self.submit.pack()

    def page_3(self):
        for i in self.master.winfo_children():
            i.destroy()
            
        self.frame3 = Frame(self.master, width=300, height=300)
        self.frame3.pack()
        self.new_file = ttk.Label(self.frame3, text='Файл создан')
        self.new_file.pack()
        self.page1_btn = ttk.Button(self.frame3, text="Вернуться в начало", command=self.page_1)
        self.page1_btn.pack()
        self.quit_btn = ttk.Button(self.frame3, text="Выход", command=self.close_app)
        self.quit_btn.pack()

    def close_app(n):
        root.destroy()

    def resize(self, event):
        region = self.canvas.bbox(tk.ALL)
        self.canvas.configure(scrollregion=region)

root = Tk()
app(root)
root.mainloop()
