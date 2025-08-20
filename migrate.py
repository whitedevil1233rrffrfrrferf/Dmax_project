import sqlite3
import mysql.connector
from mysql.connector import Error
import os

def migrate_sqlite_to_mysql(sqlite_db_path, mysql_config):
    try:
        sqlite_conn = sqlite3.connect(sqlite_db_path)
        sqlite_cursor = sqlite_conn.cursor()

        sqlite_cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
        tables = sqlite_cursor.fetchall()

        mysql_conn = mysql.connector.connect(**mysql_config)
        mysql_cursor = mysql_conn.cursor()

        for table in tables:
            table_name = table[0]
            sqlite_cursor.execute(f"PRAGMA table_info({table_name})")
            columns = sqlite_cursor.fetchall()

            columns_def = []
            pk_columns = []
            for col in columns:
                col_name = col[1]
                col_type = col[2].upper()
                is_notnull = col[3]
                is_pk = col[5]

                # SQLite to MySQL type mapping
                if col_type in ('INTEGER', 'INT'):
                    mysql_type = 'INT'
                elif col_type in ('REAL', 'FLOAT', 'DOUBLE'):
                    mysql_type = 'FLOAT'
                elif col_type == 'BOOLEAN':
                    mysql_type = 'TINYINT(1)'
                elif 'CHAR' in col_type or 'TEXT' in col_type:
                    mysql_type = 'VARCHAR(255)' if 'CHAR' in col_type else 'TEXT'
                else:
                    mysql_type = 'TEXT'

                col_def = f"`{col_name}` {mysql_type}"
                if is_notnull:
                    col_def += " NOT NULL"
                if is_pk:
                    pk_columns.append(col_name)
                    if col_name.lower() == "id" and mysql_type.startswith("INT"):
                        col_def += " AUTO_INCREMENT"

                columns_def.append(col_def)

            if pk_columns:
                columns_def.append(f"PRIMARY KEY ({', '.join(pk_columns)})")

            create_table_sql = f"CREATE TABLE IF NOT EXISTS `{table_name}` ({', '.join(columns_def)})"
            print(f"[Creating] {table_name} -> {create_table_sql}")
            mysql_cursor.execute(create_table_sql)

            sqlite_cursor.execute(f"SELECT * FROM {table_name}")
            rows = sqlite_cursor.fetchall()

            if rows:
                placeholders = ", ".join(["%s"] * len(columns))
                insert_sql = f"INSERT INTO `{table_name}` VALUES ({placeholders})"
                mysql_cursor.executemany(insert_sql, rows)

        mysql_conn.commit()
        print(f"✅ Migration successful: {sqlite_db_path}")

    except Error as e:
        print(f"❌ Error migrating {sqlite_db_path}: {e}")
    finally:
        if 'sqlite_conn' in locals():
            sqlite_conn.close()
        if 'mysql_conn' in locals():
            mysql_conn.close()


# ✅ MySQL configuration
mysql_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '1234',
    'database': 'test'  # Must exist in MySQL
}

# ✅ List all your SQLite databases (use full path if needed)
sqlite_dbs = [
    'instance/employees.db',
    'instance/dform.db',
    'instance/empinfo.db',
    'instance/opexcellence.db',
    'instance/project_targets.db',
    'instance/target_columns.db',
    'instance/dmax_approval.db'
]

# ✅ Run migration for each DB
for db in sqlite_dbs:
    if os.path.exists(db):
        print(f"\n📦 Migrating {db}...")
        migrate_sqlite_to_mysql(db, mysql_config)
    else:
        print(f"⚠️  Database file not found: {db}")
