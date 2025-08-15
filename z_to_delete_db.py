import sqlite3
conn = sqlite3.connect("results.db")
cur = conn.cursor()
cur.execute("ALTER TABLE datasets ADD COLUMN col_order TEXT")
conn.commit()
conn.close()