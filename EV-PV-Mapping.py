# -*- coding: utf-8 -*-
"""
Created on Wed Oct 14 17:35:46 2020

@author: sdhulipa
"""

###############################################################################
# Surya
# NREL
# October 2020
# Generate a CSV with mapping information from EIA-860 Solar Plants
# to EV SS data acquired recently
# Using Data extracted from .pwb files shared by Nader.
# Extracted_Gen_Data-MMWG_2020SUM_2019Series-WECC_Apr2019_20HS3ap.xlsx
###############################################################################
# Importing Required Packages
###############################################################################
import pandas as pd
from geopy import distance, Point
import numpy as np 
###############################################################################
# Location of Extracted Generator Data
###############################################################################
# Mapping to all SS extracted

# Location of EIA-860 PV and Wind Plant List

Location = r'/Users/sdhulipa/Desktop/NREL-Github/EIA860PV-Wind'


Location_3 = Location + '/EIA860/EIA860-2019ER-April2020-SolarPV-List-Updated.xlsx' 


df_pv_plants = pd.read_excel(Location_3)
df_pv_plants = df_pv_plants.drop_duplicates(subset = ['Plant ID', 'Plant State'], keep = 'first').reset_index(drop = True)

Location_5 = Location+'/NARIS to Energy Visuals mapping/New Energy Visuals Data/Nader/Energy-Visuals-SS-Data.xlsx'

df_ss_data = pd.read_excel(Location_5)
df_ss_data = df_ss_data.drop(df_ss_data.index[df_ss_data['# of Buses'] == 0])

df_pv_plants['SS Name'] = None
df_pv_plants['SS Number'] = None
df_pv_plants['Mapped EIA BA'] = None
df_pv_plants['Min Distance'] = None

for idx in df_pv_plants.index:
    
    gen_point = Point(df_pv_plants['Latitude'].loc[idx],df_pv_plants['Longitude'].loc[idx])
    
    #ba_idx = df_ss_data.index[df_ss_data['Area Name'] == df_pv_plants['Balancing Authority Code'].loc[idx]]
    print ("Solar PV Plant: %5d" %idx)
    # if (len(ba_idx) != 0):
    #     geometry = [Point(x,y) for x,y in zip(df_ss_data['Latitude'].loc[ba_idx],df_ss_data['Longitude'].loc[ba_idx])]
    #     print ("BA Match found")
    # else:
    #     geometry = [Point(x,y) for x,y in zip(df_ss_data['Latitude'],df_ss_data['Longitude'])]
    #     print ("BA Match not found")
    geometry = [Point(x,y) for x,y in zip(df_ss_data['Latitude'],df_ss_data['Longitude'])]
    gen_bus_distance =[]
    
    for bus_point in geometry:
        gen_bus_distance.append(distance.distance(gen_point,bus_point).miles)
        
    
    indices_min = [j for j, x in enumerate(gen_bus_distance) if x ==gen_bus_distance[np.argmin(gen_bus_distance)]]
    distance_min = [gen_bus_distance[k] for k in  indices_min]
   
    df_pv_plants.at[idx,'SS Name'] = df_ss_data['Sub Name'].iloc[indices_min[0]]
    df_pv_plants.at[idx,'SS Number'] = df_ss_data['Sub Num'].iloc[indices_min[0]]
    df_pv_plants.at[idx, 'Mapped EIA BA'] =df_ss_data['Area Name'].iloc[indices_min[0]]
    df_pv_plants.at[idx, 'Min Distance'] = gen_bus_distance[indices_min[0]]
    
    print ("SS Match found")

df_pv_plants.to_excel("EIA860-SolarPV-EV-SS-Mapping-New.xlsx", index=False)

zero_dis_idx_gens = df_pv_plants.index[df_pv_plants['Min Distance']==0]
interval_dis_idx_1_gens= [j for j, x in enumerate(df_pv_plants['Min Distance']) if 0<x<5]
interval_dis_idx_2_gens= [j for j, x in enumerate(df_pv_plants['Min Distance']) if 5<x<20]
interval_dis_idx_3_gens= [j for j, x in enumerate(df_pv_plants['Min Distance']) if x>20]


# 0 < Min_Distance < 5 miles - 3186
# 5 < Min_Distance < 20 miles - 362
# Min_Distance > 20 miles - 74