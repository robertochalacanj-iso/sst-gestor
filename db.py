# db.py - Capa de compatibilidad SQLite / PostgreSQL-Supabase
import os
import re
import sqlite3

DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
PGHOST = os.environ.get('PGHOST', '').strip()
USE_POSTGRES = bool(DATABASE_URL or PGHOST)

if USE_POSTGRES:
    import psycopg2
    from psycopg2 import OperationalError
    IntegrityError = psycopg2.IntegrityError
else:
    psycopg2 = None
    OperationalError = sqlite3.OperationalError
    IntegrityError = sqlite3.IntegrityError

TABLAS_CON_ID = {'clientes', 'propuestas', 'servicios', 'proyectos', 'entregables'}


def _sqlite_path():
    """SQLite local para desarrollo. En Render se recomienda Supabase/Postgres."""
    data_dir = os.environ.get('DATA_DIR', '').strip()
    if data_dir:
        os.makedirs(data_dir, exist_ok=True)
        return os.path.join(data_dir, 'sst.db')
    return os.environ.get('SQLITE_DB_PATH', 'sst.db')


class CursorPostgres:
    """Cursor con API similar a sqlite3 sobre psycopg2."""
    def __init__(self, conn):
        self._cur = conn.cursor()
        self.lastrowid = None

    def _adapt_sql(self, sql):
        sql_pg = sql.replace('?', '%s')
        # Compatibilidad básica por si algún módulo antiguo usa SQLite.
        sql_pg = re.sub(r'\bINTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT\b', 'SERIAL PRIMARY KEY', sql_pg, flags=re.I)
        return sql_pg

    def execute(self, sql, params=()):
        sql_pg = self._adapt_sql(sql)
        self.lastrowid = None
        m = re.match(r'\s*INSERT\s+INTO\s+(\w+)', sql_pg, re.IGNORECASE)
        if m and m.group(1).lower() in TABLAS_CON_ID and 'RETURNING' not in sql_pg.upper():
            sql_pg = sql_pg.rstrip().rstrip(';') + ' RETURNING id'
            self._cur.execute(sql_pg, params)
            row = self._cur.fetchone()
            self.lastrowid = row[0] if row else None
            return self
        self._cur.execute(sql_pg, params)
        return self

    def executemany(self, sql, seq):
        self._cur.executemany(self._adapt_sql(sql), seq)
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
                self._conn = psycopg2.connect(DATABASE_URL, sslmode=os.environ.get('PGSSLMODE', 'require'))
            else:
                self._conn = psycopg2.connect(
                    host=PGHOST,
                    port=os.environ.get('PGPORT', '5432'),
                    user=os.environ.get('PGUSER'),
                    password=os.environ.get('PGPASSWORD'),
                    dbname=os.environ.get('PGDATABASE', 'postgres'),
                    sslmode=os.environ.get('PGSSLMODE', 'require')
                )
        else:
            self._conn = sqlite3.connect(_sqlite_path())

    def cursor(self):
        return CursorPostgres(self._conn) if USE_POSTGRES else self._conn.cursor()

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()


def get_db():
    return Conexion()


def status():
    modo = 'PostgreSQL/Supabase' if USE_POSTGRES else 'SQLite local'
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute('SELECT 1')
        c.fetchone()
        conn.close()
        return {'ok': True, 'modo': modo, 'error': None}
    except Exception as exc:
        return {'ok': False, 'modo': modo, 'error': str(exc)}
