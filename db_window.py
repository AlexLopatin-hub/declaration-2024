import tkinter as tk
import sqlite3
from tkinter import ttk, StringVar
from tkinter.messagebox import askyesno, showerror, showinfo
import db
from edit_client_window import *


class DBWindow(tk.Toplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master)

        self.title("База клиентов")
        self.geometry("480x175+550+300")
        self.resizable(False, False)

        self.transient(master)
        self.grab_set()

        # main content
        self.listbox = tk.Listbox(self, width=60, height=10, selectmode="single")
        self.listbox.pack(side="left")
        self.scroll = ttk.Scrollbar(self, command=self.listbox.yview)
        self.scroll.pack(side="left", fill="y")
        self.listbox.config(yscrollcommand=self.scroll.set)

        self.f = ttk.Frame(self)
        self.f.pack(side="left", padx=10)
        ttk.Button(self.f, text="Изменить", command=self.edit_client).pack(fill="x")
        ttk.Button(self.f, text="Удалить", command=self.delete_client).pack(fill="x")
        ttk.Button(self.f, text="Экспорт", command=self.export_in_txt).pack(fill="x", pady=(70, 0))

        data = db.get_table()
        for item in data:
            self.listbox.insert(tk.END, item)

        self.focus_set()
        master.wait_window(self)


    def edit_client(self):
        try:
            order_in_list = self.listbox.curselection()[0]
        except IndexError:
            showerror(title="Ошибка", message="Не выбран клиент")
            return

        client_id = self.listbox.get(order_in_list)[0]
        client_name = StringVar()
        client_phone = StringVar()
        client_name.set(self.listbox.get(order_in_list)[1])
        client_phone.set(self.listbox.get(order_in_list)[2])

        EditClientWindow(self, client_name, client_phone)

        conn = db.open_connection()
        curr = conn.cursor()
        curr.execute(f"""
            UPDATE clients
            SET
                name = "{client_name.get()}",
                phone = "{client_phone.get()}"
            WHERE
                id = {client_id};
        """)
        curr.close()
        conn.commit()
        conn.close()
        print("<log> Closed database")

        self.listbox.delete(order_in_list)

        conn = sqlite3.connect("clients.db")
        curr = conn.cursor()
        curr.execute(f"SELECT * FROM clients WHERE id = {client_id};")
        client = curr.fetchall()
        curr.close()
        conn.close()
        if client:
            self.listbox.insert(order_in_list, client[0])


    def delete_client(self):
        try:
            order_in_list = self.listbox.curselection()[0]
        except IndexError:
            showerror(title="Ошибка", message="Не выбран клиент")
            return
        ans = askyesno(title="Вы уверены?", message="Вы уверены, что хотите удалить клиента?")
        if not ans:
            return

        client_id = self.listbox.get(order_in_list)[0]
        conn = db.open_connection()
        curr = conn.cursor()
        cmnd = f"DELETE FROM clients WHERE id = {client_id}"
        print(f"<log> Executing command: {cmnd}")
        curr.execute(cmnd)
        curr.close()
        conn.commit()
        conn.close()
        print("<log> Closed database")
        self.listbox.delete(order_in_list)


    def export_in_txt(self):
        conn = db.open_connection()
        curr = conn.cursor()
        curr.execute("SELECT * FROM clients;")
        data = curr.fetchall()
        curr.close()
        conn.close()
        print("<log> Closed database")
        with open(f".\\clients.txt", "a") as f:
            for client in data:
                f.write(" ".join(client[1:]) + "\n")
        showinfo(message="Результат сохранён в папке с программой.")
