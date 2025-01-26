import os
import main_process
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import askyesno, showerror, showinfo
import tempfile, base64, zlib


ICON = zlib.decompress(base64.b64decode("eJxjYGAEQgEBBiDJwZDBysAgxsDAoAHEQCEGBQaIOAg4sDIgACMUj4JRMApGwQgF/ykEAFXxQRc="))
_, ICON_PATH = tempfile.mkstemp()
with open(ICON_PATH, "wb") as icon_file: icon_file.write(ICON)
root = tk.Tk()
root.iconbitmap(default=ICON_PATH)
root.title('Выберите папку')
root.resizable(False, False)
root.geometry("400x100+750+400")

for c in range(2): root.columnconfigure(index=c, weight=1)
for r in range(3): root.rowconfigure(index=r, weight=1)


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
                if askyesno(title="Готово", message="Результат сохранён на рабочем столе. Открыть в проводнике?"):
                    os.system(f'explorer {res}')
            except FileExistsError:
                showerror(title="Ошибка", message='Папка "xml" уже существует на рабочем столе, удалите или переместите её перед тем, как начать')
            except RuntimeError:
                showerror(title="Ошибка", message="Закройте приложение Декларация 2024 перед тем как запускать программу")


entry = ttk.Entry(width=45)
entry.grid(row=0, column=0, columnspan=2)

open_button = ttk.Button(root, text="Открыть папку", command=select_folder)
open_button.grid(row=0, column=2, padx=(0, 10))

enabled = tk.IntVar()
alternate_checkbutton = ttk.Checkbutton(text="Обработать только xml", variable=enabled, width=38)
alternate_checkbutton.grid(row=1, column=0, padx=0)

start_button = ttk.Button(root, text="Начать", command=start_process)
start_button.grid(row=1, column=2, padx=(0, 10))

showinfo(title="Информация", message="При стандартном режиме работы требуется указать путь к папке с dc4 файлами. "
                                     "При работе только с xml требуется указать путь к папке с xml-файлами")

root.mainloop()
