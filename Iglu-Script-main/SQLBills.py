import pandas as pd
from sqlalchemy import create_engine, DATE
from PyPDF2 import PdfReader
import re
from env import MYSQL_USERNAME, MYSQL_PASSWORD, MYSQL_HOST, MYSQL_DATABASE

df = pd.read_excel('TelstraBills.xlsx', sheet_name='Bills')

df = df.dropna(axis=1, how='all')

engine = create_engine('mysql+mysqldb://' + MYSQL_USERNAME + ':' + MYSQL_PASSWORD + '@' + MYSQL_HOST + '/' + MYSQL_DATABASE, echo= True)

dtypes = {'MMonth': DATE}

df.to_sql('TelstraBill', engine, if_exists='replace', index=False, dtype=dtypes)