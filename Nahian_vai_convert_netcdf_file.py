# convert era5 netcdf to excel


# you need to know the centroid (lat, long) of the basin/sub-basin/area of interest before using this code
# try to follow the comments attached before each section
# this code is copyright free, a gift for you from www.linkedin.com/in/nahianmh. feel free to use/ modify this code for any purpose :D

# installing the packages
from netCDF4 import Dataset
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# read the data file
data = Dataset(r'1982-91.nc')
data

# defining units
print(data.variables.keys())
times = data.variables['time'][:]
print("Time=", times)

units = data.variables['time'].units
units = units.replace(".0", "")  # Remove .0 from units
print("unit=", units[:])

lat = data.variables['latitude'][:]
# print("lat size:", len(lat))
print("lat:", lat)
long = data.variables['longitude'][:]
print("long size:", len(long))
print("long:", long)

tp = data.variables["tp"][:]
print("tp:", tp[0])

# converting era5 time values to actual dates
era5_base_date = datetime.strptime(units, "hours since %Y-%m-%d %H:%M:%S")
era5_dates = [era5_base_date + timedelta(hours=int(hours)) for hours in times]

# gage (this generates the number of points inside the area you have selected)
n = []
c = []
t = len(times)
# print(t)
sum = 0
for i in range(t):
    print(i)
    cp = tp[i]
    # print(cp)
    n.append(cp[1][3])
    # important: inside your point of interest [x][y] here, watching the gage lat long, where x,y starts from 0,1,2,...
    # to know the numbers, you need to run the code once first
print(n)
# print("len:",len(n))

# converting time series against the values
date_time_data = []
for i in range(t):
    date_time_data.append({
        "Date": era5_dates[i],
        "tp(m)": n[i]
    })

########

# accessing the data
temperature_data = data.variables['tp']

# creating a masked array, removing masked values
masked_temperature_data = np.ma.masked_invalid(temperature_data[:])

######

# insert the expected output path including the expected file name
df = pd.DataFrame(date_time_data)
print(df)
df.to_excel(r"F:\paper\ICWFM 2023\Sara\data\tp\1982-91_rangpur.xlsx", index=False)  
# write .csv if you want .csv format

