import sqlite3


def add_client_to_db(cur, data: list) -> None:
    cur.execute("INSERT INTO clients (name, phone) VALUES (?, ?);", data)


def open_connection() -> sqlite3.Connection:
    conn = None
    try:
        conn = sqlite3.connect("clients.db")
        print("<log> Successfully connected to database")
    except sqlite3.Error as e:
        print(f"Failed to connect to database: {e}")

    curr = conn.cursor()
    curr.execute("""
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            phone TEXT NOT NULL
        )
    """)
    conn.commit()
    curr.close()
    print("<log> Closed database")
    return conn


def delete_duplicates():
    conn = sqlite3.connect("clients.db")
    curr = conn.cursor()
    curr.execute("""
        DELETE FROM clients
        WHERE rowid NOT IN (
            SELECT MIN(rowid)
            FROM clients
            GROUP BY name
        );
    """)
    curr.close()
    conn.commit()
    conn.close()

def get_table():
    conn = sqlite3.connect("clients.db")
    curr = conn.cursor()
    curr.execute("SELECT * FROM clients;")
    tables = curr.fetchall()
    curr.close()
    conn.close()
    return tables


if __name__=="__main__":
    pass