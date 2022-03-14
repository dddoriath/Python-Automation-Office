import pandas as pd
import numpy as np

df=pd.read_excel('rev_list.xlsx', index_col=0)
today = np.datetime64('today')
monday = np.datetime64('2022-03-07T00:00')
sunday = np.datetime64('2022-03-14T00:00')
monday_date = str(monday)[0:10]

minimum = df['Session End Date/Time']<sunday
df['minimum'] = np.where(minimum, df['Session End Date/Time'], sunday)

maximum = df['Session Start Date/Time']>monday
df['maximum'] = np.where(maximum, df['Session Start Date/Time'], monday)

df['overlap'] = df['minimum']-df['maximum']
df['overlap_boolean']=df['overlap'].dt.total_seconds()>0

df_lastweek = df[df['overlap_boolean'] == True]

df_lastweek['overlap_hour']=df_lastweek['overlap'].dt.total_seconds()/3600
df_report_filter = df_lastweek.iloc[:,[0,1,2,3,4,18,19,27]].sort_values(by=['Veh. ID','Session Start Date/Time'], ascending=(True,True))

filename_excel = "Weekly-Reservation-Report-%s.xlsx" % monday_date
df_report_filter.to_excel (filename_excel)
