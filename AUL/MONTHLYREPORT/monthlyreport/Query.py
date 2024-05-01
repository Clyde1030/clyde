### This script interacts with SQL Server to execute the SQL Codes for dealers' monthly claim reports.
### The SQL code is saved as below removing the parameter declaration on the top.
### X:\Dept.Risk.Management\Claims\SQL Claims Detail with GWR_AULDATAMART_with Post Period_03-07-2020 (updated).sql
### The entire module is imported into ReportRunner.py

import re
import logging
from calendar import monthrange 

import pyodbc
import pandas as pd

LOGGER = logging.getLogger(__name__)

# Conncet to SQL Server with PyODBC. Input are the server and database that are connected to.
def connect(server,database):
    DRIVER = '{ODBC Driver 17 for SQL Server}'    
    CONN_STRING = """
        Driver={driver};
        Server={server};
        Database={database};
        Trusted_Connection=yes;""".format(driver=DRIVER,server=server,database=database)
    conn = pyodbc.connect(CONN_STRING)
    LOGGER.debug('Connectted with: {}'.format(CONN_STRING))
    csr = conn.cursor()
    return conn, csr

# Build the sql headers according to the roles(dealer, dealer group, financial company) the query is executing on.
def build_sql_headers(role,id,year,month,mode):
    '''
    role: 
        "fc" for finance company/lenders,\n
        "dlr" for dealer,\n
        "dlr_group" for dealer group\n
    mode default to s: 
        "s" for single month,\n
        "i" for ITD\n
    '''
    last_day = monthrange(int(year),int(month))[1]
    asofperiod = year+month.zfill(2)
    fromdate = str(month)+'/01/'+str(year)    
    todate = str(month)+'/'+str(last_day)+'/'+str(year)    


    header = """DECLARE @TO_DATE_PAID VARCHAR(MAX) = '{td}';\nDECLARE @FROM_SOLD_DATE VARCHAR(MAX) = NULL;\nDECLARE @TO_SOLD_DATE VARCHAR(MAX) = NULL;\nDECLARE @ASOF_PERIOD CHAR(6) = '{aop}';\nDECLARE @STATUS VARCHAR(MAX) = 'Paid';\n""".format(td=todate,aop=asofperiod)
    
    if mode == 'i':
        header = "DECLARE @FROM_DATE_PAID VARCHAR(MAX) = NULL;\n"+header
    else:
        header = "DECLARE @FROM_DATE_PAID VARCHAR(MAX) = '{}';\n".format(fromdate)+header

    if role == "fc":
        header = header + "DECLARE @FC_ID VARCHAR(MAX) = {};\nDECLARE @DEALER_ID VARCHAR(MAX) = NULL;\nDECLARE @DEALER_GROUP VARCHAR(MAX) = NULL;\n".format(id)
    elif role == "dlr":
        header = header + "DECLARE @FC_ID VARCHAR(MAX) = NULL;\nDECLARE @DEALER_ID VARCHAR(MAX) = {};\nDECLARE @DEALER_GROUP VARCHAR(MAX) = NULL;\n".format(id)
    elif role == "dlr_group":
        header = header + "DECLARE @FC_ID VARCHAR(MAX) = NULL;\nDECLARE @DEALER_ID VARCHAR(MAX) = NULL;\nDECLARE @DEALER_GROUP VARCHAR(MAX) = {};\n".format(id)
    else:
        raise RuntimeError("Unrecognized role " + role)

    header = header + "DECLARE @ADMINISTRATOR VARCHAR(MAX) = NULL;"
    LOGGER.debug(header)

    return header

# Execute the query and return the results as a Pandas dataframe and the field information
def query(qry_text, server=None, database=None):
    '''Execute the sql text string and return the query result as a list using PyODBC'''
    cn, csr = connect(server,database)

    # "GO" in the query would fail the result
    go_parse_regex = re.compile(r'^s*[Gg][Oo]\s*$', re.MULTILINE)
    script_with_proper_go = re.sub(go_parse_regex, 'GO', qry_text)
    batches = script_with_proper_go.split('GO\n')
    
    # The solution is to split the query into different batches. Execute by batches split by "GO"
    for b in batches:
        trimmed = re.sub(r'^\s*$', '', b)
        if trimmed:
            try:
                csr.execute(trimmed)
                if csr.rowcount:
                    try:
                        rows = csr.fetchall()
                        cols = [column[0] for column in csr.description]
                        df = pd.DataFrame.from_records(data=rows,columns=cols)
                        row = df.shape[0]
                        LOGGER.info('{} records returned from query.'.format(row))
                    except pyodbc.ProgrammingError:
                        pass

                while csr.nextset(): 
                    if csr.rowcount:
                        try:
                            rows = csr.fetchall()
                            cols = [column[0] for column in csr.description]
                            df = pd.DataFrame.from_records(data=rows,columns=cols)
                            row = df.shape[0]
                            LOGGER.info('{} records returned from query.'.format(row))
                        except pyodbc.ProgrammingError:
                            pass
            except:
                LOGGER.warning(trimmed)
                raise

    LOGGER.debug('PyODBC Connection closed.')
    cn.close()

    return df, cols 

