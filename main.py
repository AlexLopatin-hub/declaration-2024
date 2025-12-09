import os
import main_process
import db
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import askyesno, showerror, showinfo
import tempfile, base64, zlib
from info_window import *
from db_window import *


ICON = zlib.decompress(base64.b64decode("eJxjYGAEQgEBBiDJwZDBysAgxsDAoAHEQCEGBQaIOAg4sDIgACMUj4JRMApGwQgF/ykEAFXxQRc="))
_, ICON_PATH = tempfile.mkstemp()
with open(ICON_PATH, "wb") as icon_file: icon_file.write(ICON)
root = tk.Tk()
root.iconbitmap(default=ICON_PATH)
root.title('Выберите папку')
root.resizable(False, False)
root.geometry("450x100+750+400")

for r in range(3): root.rowconfigure(index=r, weight=1)
for c in range(3): root.columnconfigure(index=c, weight=1)


def select_folder():
    folder = fd.askdirectory()
    if folder:
        entry.delete(0, 'end')
        entry.insert(0, folder)


def start_process():
    folder = entry.get().replace("/", "\\")
    if folder == "":
        showerror(title="Ошибка", message="Не указана директория")
    else:
        ans = askyesno(title="Вы уверены?", message="Вы уверены что хотите начать? Процесс будет невозможно прервать")
        if ans: ans = askyesno(title="Вы уверены?", message="Вы уверены что хотите начать? Процесс будет невозможно прервать")
        if ans:
            try:
                res = main_process.main(folder, enabled.get())
                showinfo(message='Готово. Результат сохранён в clients.db в папке с программой. '
                                 'Сделать экспорт в текстовый файл можно в окне "База клиентов"')
            except FileExistsError:
                showerror(title="Ошибка", message='Не удалось создать папку с названием "xml" в каталоге "C:\\", удалите или переместите её перед тем, как начать')
            except RuntimeError:
                showerror(title="Ошибка", message="Закройте приложение Декларация 2024 перед тем как запускать программу")

def show_info():
    window = InfoWindow(root, enabled.get())


def list_database():
    window = DBWindow(root)


if __name__=="__main__":
    entry = ttk.Entry(width=55)
    entry.grid(row=0, column=0, columnspan=2, padx=10)

    open_button = ttk.Button(root, text="Открыть папку", command=select_folder)
    open_button.grid(row=0, column=2, padx=(0, 10), ipadx=10)

    enabled = tk.IntVar()
    alternate_checkbutton = ttk.Checkbutton(text="Обработать только xml", variable=enabled, width=38)
    alternate_checkbutton.grid(row=1, column=0, padx=(10, 0))

    start_button = ttk.Button(root, text="Начать", command=start_process)
    start_button.grid(row=1, column=2, padx=(0, 10), ipadx=10)

    info_button = ttk.Button(root, text="Инструкция", command=show_info)
    info_button.grid(row=2, column=2, padx=(0, 10), ipadx=10)

    list_button = ttk.Button(root, text="База клиентов", command=list_database)
    list_button.grid(row=2, column=1, padx=(0, 10), ipadx=10)

    root.mainloop()
