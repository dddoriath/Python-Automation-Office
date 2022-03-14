import pandas as pd
import numpy as np
import time

today = np.datetime64('today')

df=pd.read_excel('rev_list.xlsx', index_col=0)

df['today']=(df['Session Start Date/Time']<=today+np.timedelta64(24, 'h')) & (df['Session End Date/Time']>=today-np.timedelta64(24, 'h')) & (df['Approval Status'] == 'Approved')

df_today = df[df['today'] == True]

df_today_filter = df_today.iloc[:,[0,1,2,3,4,5,]].sort_values(by=['Veh. ID','Session Start Date/Time'], ascending=(True,True))

filename_excel = "Reservation-%s.xlsx" % today

df_today_filter.to_excel (filename_excel)
