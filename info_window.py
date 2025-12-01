import tkinter as tk
from tkinter import ttk


class InfoWindow(tk.Toplevel):
    def __init__(self, master,  mode: int, **kwargs):
        super().__init__(master)

        self.title("Инструкция")
        self.geometry("400x130+700+350")
        self.resizable(False, False)

        self.transient(master)
        self.grab_set()

        if not mode:
            info = ("При стандартном режиме работы требуется указать путь к папке с dc4 файлами. Программа извлечёт "
                    "информацию из них и запишет её в xml файлы в папке на диске C, а затем соберёт всё в единый "
                    ".txt файл в той же папке.")
        else:
            info = ("При работе только с xml требуется указать путь к папке с уже извлечёнными xml-файлами."
                    "Программа пропустит первый этап и перейдёт сразу к сбору всей информации о клиентах в "
                    "единую таблицу.")

        self.info = ttk.Label(self, text=info, wraplength=380, padding=10)
        self.info.pack(expand=True, fill='both')

        close_btn = ttk.Button(self, text="Закрыть", command=self.destroy)
        close_btn.pack(pady=10)

        self.focus_set()
        master.wait_window(self)
