import sqlite3

def connect():
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("""
                CREATE TABLE IF NOT EXISTS members (
                    id INTEGER PRIMARY KEY,
                    memID TEXT,
                    name TEXT,
                    gender TEXT,
                    dob TEXT,
                    baptised TEXT,
                    residence TEXT,
                    department TEXT,
                    tel TEXT,
                    nationality TEXT,
                    email TEXT
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS member_images (
                    id INTEGER PRIMARY KEY,
                    member_id INTEGER,
                    image_data BLOB,
                    FOREIGN KEY (member_id) REFERENCES members(id)
                )
            """)
            con.commit()
    except sqlite3.Error as e:
        print("Error creating tables:", e)

def insert_member(memID="", name="", gender="", dob="", baptised="", residence="", department="", tel="", nationality="", email=""):
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("""
                INSERT INTO members 
                VALUES (NULL,?,?,?,?,?,?,?,?,?,?)
            """, (memID, name, gender, dob, baptised, residence, department, tel, nationality, email))
            member_id = cur.lastrowid  # Get the ID of the newly inserted member
            con.commit()
        return member_id
    except sqlite3.Error as e:
        print("Error inserting member:", e)
        return None

def insert_image_data(member_id, image_data):
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("INSERT INTO member_images(member_id, image_data) VALUES (?, ?)", (member_id, image_data))
            con.commit()
    except sqlite3.Error as e:
        print("Error inserting image data:", e)

def view_members():
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("SELECT * FROM members")
            rows = cur.fetchall()
        return rows
    except sqlite3.Error as e:
        print("Error viewing members:", e)
        return None

def view_member_with_image(member_id):
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("""
                SELECT m.*, i.image_data AS member_image
                FROM members m
                LEFT JOIN member_images i ON m.id = i.member_id
                WHERE m.id = ?
            """, (member_id,))
            row = cur.fetchone()
        return row
    except sqlite3.Error as e:
        print("Error viewing member with image:", e)
        return None

def delete_member(member_id):
    try:
        with sqlite3.connect("database.db") as con:
            cur = con.cursor()
            cur.execute("DELETE FROM members WHERE id = ?", (member_id,))
            con.commit()
    except sqlite3.Error as e:
        print("Error deleting member:", e)

connect()
