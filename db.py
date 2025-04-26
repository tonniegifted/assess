import mysql.connector
db=mysql.connector.connect(host="localhost",
                           user="root",
                           password="print",
                           database="remedial")
cur=db.cursor()

        
        
def dbfunc():

    cur.execute("SELECT subject_score FROM score")
    for i in cur.fetchall():
        print("subject scores",i[0])
    cur.execute("SELECT grandtotal FROM grand")
    for j in cur.fetchall():
        print("grand total",j[0])
    cur.execute("SELECT total_score FROM total")
    for k in cur.fetchall():
        print("total",k[0])
# dbfunc()

def deleting():
    cur.execute("""
                DELETE FROM score
                """)
    cur.execute("""
                DELETE FROM grand
                """)
    cur.execute("""
                DELETE FROM total
                """)
    db.commit()
    print("Records deleted successfully!")
# deleting()

