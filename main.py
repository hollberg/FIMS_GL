"""main.py

Code to run GL Account Changes in CFGA FIMS
"""

# Mitch Scripts
# Standard Library
import time
# External Libraries
# Local Modules
import scripts
import sqls

time.sleep(5)

# Query GL Account/Year to Map
fims_conn = sqls.make_fims_connection()
fims_cursor = fims_conn.cursor()

sqlite_conn = sqls.make_sqlite_conn(sqls.SQLITE_DIRECTORY, sqls.SQLITE_FILE)
sqlite_cursor = sqlite_conn.cursor()
sqlite_write_cursor = sqlite_conn.cursor()

sqlite_cursor.execute("""SELECT * FROM coa_sqlite LIMIT 1""")
cols = [sqlite_cursor.description[i][0]
        for i in range(0,len(sqlite_cursor.description))]

sqlite_cursor.execute(f"""
SELECT {','.join(cols)}
FROM coa_sqlite
WHERE complete = 0 
    and GL_Yr <= 2011
-- LIMIT 1;
""")
for result in sqlite_cursor:
    # result = sqlite_cursor.fetchone()
    print(result)
    moo = 'boo'

    # Run update Here

    time_begin = time.time()

    time.sleep(5)

    scripts.map_gl(fisc_yr=result.GL_Yr,
                   gl_old=result.Acct_old,
                   gl_new=result.Acct_new)

    time_end = time.time()
    time_elapsed = time_end-time_begin

    elapsed = time.time() - time_begin
    print(f'*** mapping {result.GL_Yr} account {result.Acct_old} to {result.Acct_new} took {elapsed/60} minutes')

    # Get Credit/Debit Balance for "New" GL Account
    gl_figures = fims_cursor.execute(sqls.gl_acct_sum,
                                     [int(result.GL_Yr), int(result.Acct_new)]).fetchone()

    print(gl_figures)

    # Compare values:
    delta_debits = result.Debit_total - float(gl_figures[2])
    delta_credits = result.Credit_total - float(gl_figures[3])
    print(f'Delta debits is: {delta_debits}, Delta Credits is: {delta_credits}')
    doesnt_tie = 0
    if delta_credits != 0 or delta_debits != 0:
        doesnt_tie = 1

    # Update SQLite
    # Update row to indicate completion


    sql_update = f"""
    UPDATE coa_sqlite
    SET time_begin = {time_begin},
        time_end = {time_end},
        complete = 1,
        IsError = {doesnt_tie},
        Debit_new = {float(gl_figures[2])},
        Credit_new = {float(gl_figures[3])}
        
    WHERE Acct_old = {result.Acct_old} AND
          Acct_new = {result.Acct_new} AND
          GL_Yr = {result.GL_Yr};"""

    sqlite_write_cursor.execute(sql_update)
    sqlite_conn.commit()