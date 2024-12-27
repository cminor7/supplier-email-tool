# internal libraries
from getpass import getuser
from os import path

# external libraries
import snowflake.connector # pip install snowflake-connector-python
import pandas as pd # pip install pandas


def sfConn():
    conn = snowflake.connector.connect(
    user=getuser(),
    account='wwgrainger.us-east-1',
    password='',
    authenticator="externalbrowser",
    token='',
    warehouse='GSCCE_ANALYTICS_S',
    database='',
    schema='',
    role='ISP_SVC'
    )
    return conn


def getEmail():
	racfid = getuser().upper()
	query = f"""SELECT WORK_EMAIL FROM PUBLISH.GSCCE.EMPLOYEE_EDV WHERE USERNAME= '{racfid}'"""
	df = pd.read_sql(query, sfConn())
	email = df['WORK_EMAIL'].iloc[0]

	with open(f'{racfid}_email.txt', 'w') as text_file:
		text_file.write(email)


racfid = getuser().upper()
if not path.exists(f'{racfid}_email.txt'):
    getEmail()