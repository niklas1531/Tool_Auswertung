from tkinter import Tk, ttk
from tkinter import filedialog
from tool_template import CreateTemplate

import tkinter as tk





class Gui:
    def __init__(self, master):
        self.master = master
        master.title("Create Indiv. Template")
        self.path = ''

        self.frame = ttk.Frame(root, padding=10)
        self.frame.grid(row=0, column=0)

        # Feld mit Path
        self.file_name_var = tk.StringVar()
        self.file_name_label = ttk.Label(root, textvariable=self.file_name_var)
        self.file_name_label.grid(row=1, column=0)
        # Button zur Auswahl der Datei
        self.select_file_button = ttk.Button(self.frame, text="Datei auswählen", command=self.open_file)
        self.select_file_button.grid(row=3, column=1)

        # Button zur Verarbeitung der Datei
        self.process_file_button = ttk.Button(self.frame, text="Datei verarbeiten", command=self.process_file)
        self.process_file_button.grid(row=4, column=1)

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        print(path)
        self.file_name_var.set(path)
        self.path = path
    def process_file(self):
        createTemplate_obj = CreateTemplate()
        createTemplate_obj.createTemplate(self.path)
        root.quit()

root = Tk()
w = 300
h = 120
ws = root.winfo_screenwidth()
hs = root.winfo_screenheight()
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
my_gui = Gui(root)
root.mainloop()
