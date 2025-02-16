# internal libraries
from getpass import getuser
from os import path, makedirs

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
	query = f"""SELECT WORK_EMAIL FROM PUBLISH.GSCCE.EMPLOYEE_EDV WHERE USERNAME = '{racfid}'"""
	df = pd.read_sql(query, sfConn())
	email = df['WORK_EMAIL'].iloc[0]

	with open('user_email.txt', 'w') as text_file:
		text_file.write(email)


def getSecret():
    query = """SELECT TO_VARCHAR(DECRYPT(API_KEY, '-sAlly$N3verLies!'), 'UTF-8') AS API_KEY FROM ISP.RA.SMTP_SECRET"""
    df = pd.read_sql(query, sfConn())
    api_key = df['API_KEY'].iloc[0]

    with open('DEVLOPER_FILES/api_key.txt', 'w') as text_file:
        text_file.write(api_key)


if not path.exists('user_email.txt'):
      getEmail()

if not path.exists('DEVLOPER_FILES/api_key.txt'):
      getSecret()

if not path.exists('ATTACHMENT'):
    makedirs('ATTACHMENT')