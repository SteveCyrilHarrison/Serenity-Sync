import sqlite3

def connect():
    try:
        conn = sqlite3.connect('attendance_database.db')
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                program TEXT,
                male INTEGER,
                female INTEGER,
                youth INTEGER,
                child INTEGER,
                total INTEGER
            )
        ''')
        conn.commit()
    except sqlite3.Error as e:
        print("Error creating table:", e)
    finally:
        conn.close()

def insert_attend(date,program,male,female,youth,child,total):
    try:
        conn = sqlite3.connect('attendance_database.db')
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO attendance (date,program,male,female,youth,child,total)
            VALUES (?,?,?,?,?,?,?)
        ''', (date, program, male,female,youth,child,total))
        conn.commit()
    except sqlite3.Error as e:
        print("Error inserting data:", e)
    finally:
        conn.close()

def view_attend():
    try:
        conn = sqlite3.connect('attendance_database.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM attendance")
        rows = cursor.fetchall()
        return rows
    except sqlite3.Error as e:
        print("Error viewing data:", e)
    finally:
        conn.close()

def delete_attend(id):
    try:
        conn = sqlite3.connect('attendance_database.db')
        cursor = conn.cursor()
        cursor.execute("DELETE FROM attendance WHERE id = ?", (id,))
        conn.commit()
    except sqlite3.Error as e:
        print("Error deleting data:", e)
    finally:
        conn.close()

# Create the table if not exists
connect()
