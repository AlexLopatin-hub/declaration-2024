import tkinter as tk
from tkinter import ttk


class EditClientWindow(tk.Toplevel):
    def __init__(self, master, name: tk.StringVar, phone: tk.StringVar, **kwargs):
        super().__init__(master)

        self.title("Введите данные клиента")
        self.geometry("400x150+700+350")
        self.resizable(False, False)

        self.transient(master)
        self.grab_set()

        self.name = name
        self.phone = phone

        ttk.Label(self, text="Имя").pack(anchor='w', padx=20)
        self.new_name = ttk.Entry(self)
        self.new_name.pack(fill="x", padx=20)
        self.new_name.insert(0, name.get())

        ttk.Label(self, text="Телефон").pack(anchor='w', padx=20, pady=(10, 0))
        self.new_phone = ttk.Entry(self)
        self.new_phone.pack(fill="x", padx=20)
        self.new_phone.insert(0, phone.get())

        f = ttk.Frame(self)
        f.pack(anchor="n", pady=(20, 0))
        ttk.Button(f, text="Подтвердить", command=self.confirm).pack(side="right", padx=5)
        ttk.Button(f, text="Закрыть", command=self.destroy).pack(side="left", padx=5)

        self.focus_set()
        master.wait_window(self)


    def confirm(self):
        self.name.set(self.new_name.get())
        self.phone.set(self.new_phone.get())
        self.destroy()