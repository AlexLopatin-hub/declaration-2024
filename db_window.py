import tkinter as tk
from tkinter import ttk

from gevent.testing.travis import command

import db

class DBWindow(tk.Toplevel):
    def __init__(self, master, **kwargs):
        super().__init__(master)

        self.title("База клиентов")
        self.geometry("420x175+550+300")
        self.resizable(False, False)

        self.transient(master)
        self.grab_set()

        # main content
        self.listbox = tk.Listbox(self, width=50, height=10, selectmode="single")
        self.listbox.pack(side="left")
        self.scroll = ttk.Scrollbar(self, command=self.listbox.yview)
        self.scroll.pack(side="left", fill="y")
        self.listbox.config(yscrollcommand=self.scroll.set)

        self.f = ttk.Frame(self)
        self.f.pack(side="left", padx=10)
        ttk.Button(self.f, text="Изменить", command=self.foo).pack(fill="x")
        ttk.Button(self.f, text="Удалить", command=self.delete_client).pack(fill="x")
        ttk.Button(self.f, text="Экспорт", command=self.export_in_txt).pack(fill="x", pady=(70, 0))

        data = db.get_table()
        for item in data:
            self.listbox.insert(tk.END, item)

        self.focus_set()
        master.wait_window(self)


    def foo(self):
        selection = self.listbox.curselection()[0]
        print(selection)

    def edit_client(self, name: str, new_info: list):
        conn = db.open_connection()
        curr = conn.cursor()
        curr.execute(f"""
            UPDATE clients
            SET
                name = {new_info[0]}
                phone = {new_info[1]}
            WHERE
                name = {name};
        """)
        curr.close()
        conn.close()
        print("<log> Closed database")


    def delete_client(self):
        order_in_list = self.listbox.curselection()[0]
        client_id = self.listbox.get(order_in_list)[0]
        conn = db.open_connection()
        curr = conn.cursor()
        cmnd = f"DELETE FROM clients WHERE id = {client_id}"
        print(f"<log> Executing command: {cmnd}")
        curr.execute(cmnd)
        conn.commit()
        curr.close()
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
