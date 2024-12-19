import sqlite3

def connect():
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("""CREATE TABLE IF NOT EXISTS  (
            id INTEGER PRIMARY KEY,
            offtype TEXT,
            amount TEXT,
            date TEXT,
            service TEXT
        )""")

        cur.execute("""CREATE TABLE IF NOT EXISTS tithe (
            id INTEGER PRIMARY KEY,
            memID TEXT,
            name TEXT,
            date TEXT,
            amount TEXT,
            contact TEXT
        )""")
        con.commit()
    except sqlite3.Error as e:
        print("Error creating tables:", e)
    finally:
        con.close()

def GiveInsert(offtype="", amount="", date="", service=""):
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("""
            INSERT INTO  VALUES(NULL,?,?,?,?)
        """, (offtype, amount, date, service))
        con.commit()
    except sqlite3.Error as e:
        print("Error inserting data into :", e)
    finally:
        con.close()

def TitheInsert(memID="", name="", date="", amount="", contact=""):
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("""
            INSERT INTO tithe VALUES(NULL,?,?,?,?,?)
        """, (memID, name, date, amount, contact))
        con.commit()
    except sqlite3.Error as e:
        print("Error inserting data into tithe:", e)
    finally:
        con.close()

def view_Give():
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("SELECT * FROM ")
        rows = cur.fetchall()
        return rows
    except sqlite3.Error as e:
        print("Error viewing data in :", e)
    finally:
        con.close()

def view_Tithe():
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("SELECT * FROM tithe")
        rows = cur.fetchall()
        return rows
    except sqlite3.Error as e:
        print("Error viewing data in tithe:", e)
    finally:
        con.close()

def delete_Give(id):
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("DELETE FROM  WHERE id = ?", (id,))
        con.commit()
    except sqlite3.Error as e:
        print("Error deleting data from :", e)
    finally:
        con.close()

def delete_Tithe(id):
    try:
        con = sqlite3.connect("GiveDB.db")
        cur = con.cursor()
        cur.execute("DELETE FROM tithe WHERE id = ?", (id,))
        con.commit()
    except sqlite3.Error as e:
        print("Error deleting data from tithe:", e)
    finally:
        con.close()

# Create tables if not exists
connect()
