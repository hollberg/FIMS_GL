import sqlite3
import openpyxl
import pandas as pd
import pyodbc
import csv
from collections import namedtuple

SQLITE_DIRECTORY = r'N:\FIS\COA'
SQLITE_FILE = 'sqlite_convert_2009_2016V2.sqlite'

def make_fims_connection():
    conn = pyodbc.connect(
        r'DRIVER={Progress OpenEdge 10.2A driver};'
        r'DSN=Foundsql92;'
        r'HOST=app-svr2-cfga;'
        r'UID=user_id;'
        r'PWD=******;'
        r'PORT=2600;'
        r'DB=found;'
    )
    return conn


def build_sqlite():
    """
    Build and populate a SQLite database with GL Accounts to remap
    :return:
    """
    SQLITE_FILE = 'sqlite_convert_2009_2016.sqlite'
    DIRECTORY = r'N:\FIS\COA'

    sqlite_conn = sqlite3.connect(DIRECTORY + '\\' + SQLITE_FILE)
    c = sqlite_conn.cursor()

    c.execute("""
    DROP TABLE IF EXISTS 'coa_sqlite';
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS `coa_sqlite` (
    	`Acct_old`	TEXT NOT NULL,
    	`Acct_new`	TEXT NOT NULL,
    	`GL_Yr`	TEXT NOT NULL,
    	`NewName`	TEXT,
    	`Notes`	TEXT,
    	`Debit_total`	NUMERIC,
    	`Credit_total`	NUMERIC,
    	`time_begin`	REAL,
    	`time_end`	REAL,
    	`time_elapsed`	REAL,
    	`complete`	INTEGER DEFAULT 0 CHECK (complete IN (0,1)),
    	`Debit_new`	NUMERIC,
    	`Credit_new`  NUMERIC,
    	`IsError` INTEGER DEFAULT 0 CHECK (IsError IN (0,1)),
    	PRIMARY KEY ('Acct_old', 'Acct_new','GL_Yr'),
    	CHECK (complete in (0,1)));
    """)

    # Read data to pandas
    indexes = ['Acct_old', 'GL_Yr']
    je_df = pd.read_csv(r'N:\FIS\COA\coa_sqlite_import_2009_2016.csv',
                        header=0,
                        encoding='ISO-8859-1')

    # je_df = je_df.set_index(indexes)

    je_df.to_sql(name='coa_sqlite', con=sqlite_conn, if_exists='append',
                 index=False)

    sqlite_conn.commit()
    sqlite_conn.close()

def namedtuple_factory(cursor, row):
    """Returns sqlite rows as named tuples."""
    fields = [col[0] for col in cursor.description]
    Row = namedtuple("Row", fields)
    return Row(*row)


def make_sqlite_conn(sqlite_directory, sqlite_db):
    """
    :param sqlite_directory: Directory where db is stored
    :param sqlite_db: filename of the sqlite3 db
    :return: a SQLite3 database cursor object
    """
    sqlite_conn = sqlite3.connect(sqlite_directory + '\\' + sqlite_db)
    sqlite_conn.row_factory = namedtuple_factory
    return sqlite_conn

def get_conversion_data_from_sqlite():
    # todo
    pass


#   **** SQL Query Strings  ***

je_qry = """
SELECT FundId, GLYear, 
JrnlNo, EntryDate, 
ApplyDate, Descr, 
JrnlKey,  "Ref-no",
sg1, sg2, 
sg3, sg4, 
sg5, sg6, 
Debit, Credit
FROM FOUND.pub.GLJourHis
WHERE GLYear=2017 
  AND FundId='AA-SCH'
  AND sg6='11007'
"""

je_2 = """
SELECT GLYear, 
JrnlNo, EntryDate, 
ApplyDate, Descr, 
JrnlKey,  "Ref-no",
sg1, sg2, 
sg3, sg4, 
sg5, sg6, 
Debit, Credit
FROM FOUND.pub.GLJourHis
WHERE GLYear=? 
  AND sg6=?
"""

je_sums = """
SELECT GLJourHis.GLYear
	,sg6
	,Sum(Debit) AS 'Sum of Debit'
	,Sum(Credit) AS 'Sum of Credit'
	,Count(JrnlNo) AS 'Count of JrnlNo'
FROM FOUND.PUB.GLJourHis GLJourHis
GROUP BY GLYear
	,sg6
HAVING (GLYear >= 2009)
ORDER BY GLYear
	,sg6"""

gl_acct_sum = """
SELECT GLYear
	,sg6
	,Sum(Debit) AS 'Sum of Debit'
	,Sum(Credit) AS 'Sum of Credit'
	,Count(JrnlNo) AS 'Count of JrnlNo'
FROM FOUND.PUB.GLJourHis
GROUP BY GLYear
	,sg6
HAVING (GLYear = ?)
	AND (sg6 = ?)
ORDER BY GLYear
	,sg6
"""

# qry_next = f"""
# SELECT {','.join(cols)}
# FROM coa_sqlite
# WHERE complete = 0
# LIMIT 1;
# """

# build_sqlite()    # Build and populate new database