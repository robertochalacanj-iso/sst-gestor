# db.py - Capa de compatibilidad SQLite / PostgreSQL
import os
import re
import sqlite3

DATABASE_URL = os.environ.get('DATABASE_URL', '')
PGHOST = os.environ.get('PGHOST', '')
USE_POSTGRES = bool(DATABASE_URL or PGHOST)

if USE_POSTGRES:
    import psycopg2
    IntegrityError = psycopg2.IntegrityError
else:
    psycopg2 = None
    IntegrityError = sqlite3.IntegrityError

TABLAS_CON_ID = {'clientes', 'propuestas', 'servicios', 'proyectos', 'entregables'}


class CursorPostgres:
    """Cursor con la misma API de sqlite3 sobre psycopg2"""
    def __init__(self, conn):
        self._cur = conn.cursor()
        self.lastrowid = None

    def execute(self, sql, params=()):
        sql_pg = sql.replace('?', '%s')
        m = re.match(r'\s*INSERT\s+INTO\s+(\w+)', sql_pg, re.IGNORECASE)
        if m and m.group(1) in TABLAS_CON_ID and 'RETURNING' not in sql_pg.upper():
            sql_pg = sql_pg.rstrip().rstrip(';') + ' RETURNING id'
            self._cur.execute(sql_pg, params)
            row = self._cur.fetchone()
            self.lastrowid = row[0] if row else None
            return self
        self._cur.execute(sql_pg, params)
        return self

    def executemany(self, sql, seq):
        self._cur.executemany(sql.replace('?', '%s'), seq)
        return self

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    def close(self):
        self._cur.close()


class Conexion:
    def __init__(self):
        if USE_POSTGRES:
            if DATABASE_URL:
                self._conn = psycopg2.connect(DATABASE_URL, sslmode='require')
            else:
                self._conn = psycopg2.connect(
                    host=PGHOST,
                    port=os.environ.get('PGPORT', '5432'),
                    user=os.environ.get('PGUSER'),
                    password=os.environ.get('PGPASSWORD'),
                    dbname=os.environ.get('PGDATABASE', 'postgres'),
                    sslmode='require'
                )
        else:
            self._conn = sqlite3.connect('sst.db')

    def cursor(self):
        return CursorPostgres(self._conn) if USE_POSTGRES else self._conn.cursor()

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def get_db():
    return Conexion()