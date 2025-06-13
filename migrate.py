import sqlite3
import mysql.connector
from mysql.connector import Error

def migrate_sqlite_to_mysql(sqlite_db_path, mysql_config):
    try:
        # Connect to SQLite
        sqlite_conn = sqlite3.connect(sqlite_db_path)
        sqlite_cursor = sqlite_conn.cursor()

        # Get all tables
        sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = sqlite_cursor.fetchall()

        # Connect to MySQL
        mysql_conn = mysql.connector.connect(**mysql_config)
        mysql_cursor = mysql_conn.cursor()

        # Migrate each table
        for table in tables:
            table_name = table[0]
            
            # Get table schema
            sqlite_cursor.execute(f"PRAGMA table_info({table_name})")
            columns = sqlite_cursor.fetchall()
            
            # Create table in MySQL
            columns_def = ", ".join([f"`{col[1]}` TEXT" for col in columns])
            create_table_sql = f"CREATE TABLE IF NOT EXISTS `{table_name}` ({columns_def})"
            mysql_cursor.execute(create_table_sql)
            
            # Get data from SQLite
            sqlite_cursor.execute(f"SELECT * FROM {table_name}")
            rows = sqlite_cursor.fetchall()
            
            # Insert data into MySQL
            if rows:
                placeholders = ", ".join(["%s"] * len(columns))
                insert_sql = f"INSERT INTO `{table_name}` VALUES ({placeholders})"
                mysql_cursor.executemany(insert_sql, rows)

        mysql_conn.commit()
        print(f"Successfully migrated {sqlite_db_path}")

    except Error as e:
        print(f"Error: {e}")
    finally:
        if 'sqlite_conn' in locals():
            sqlite_conn.close()
        if 'mysql_conn' in locals():
            mysql_conn.close()

mysql_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '1234',
    'database': 'test'
}

# List of your SQLite databases
sqlite_dbs = [
    'instance/dform.db',
    'instance/dmax_approval.db',
    'instance/empinfo.db',
    'instance/employees.db',
    'instance/opexcellence.db',
    'instance/project_targets.db',
    'instance/target_columns.db'
]

# Migrate each database
for db in sqlite_dbs:
    migrate_sqlite_to_mysql(db, mysql_config)    