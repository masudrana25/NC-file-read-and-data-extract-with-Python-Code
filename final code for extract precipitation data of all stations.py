from netCDF4 import Dataset
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings('ignore')

# read the data file
data = Dataset(r'data_file.nc')

times = data.variables['time'][:] 
units = data.variables['time'].units
units = units.replace(".0", "")
lat = data.variables['latitude'][:]
long = data.variables['longitude'][:]
tp = data.variables["tp"][:]

era5_base_date = datetime.strptime(units, "hours since %Y-%m-%d %H:%M:%S")
era5_dates = [era5_base_date + timedelta(hours=int(hours)) for hours in times]

# Loop through each latitude and longitude pair
for lat_index in range(len(lat)):
    for lon_index in range(len(long)):
        n = []
        date_time_data = []

        for i in range(len(times)):
            cp = tp[i]
            n.append(cp[lat_index][lon_index])

            date_time_data.append({
                "Date": era5_dates[i],
                f"tp(m)_lat({lat[lat_index]})_long({long[lon_index]})": n[i]
            })

        df = pd.DataFrame(date_time_data)
        filename = f"lat({lat[lat_index]})_long({long[lon_index]}).xlsx"
        df.to_excel(filename, index=False)
        print(f"Excel file '{filename}' created.")
