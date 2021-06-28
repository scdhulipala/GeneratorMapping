#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 14:27:14 2020

@author: sdhulipa
"""


###############################################################################
# Surya
# NREL
# November 2020
# Mapping for EIA-860 Solar/Wind Plant for BA's
###############################################################################
# Importing Required Packages
###############################################################################
import pandas as pd
from geopy import distance, Point
import numpy as np 
###############################################################################
Location = r'/Users/sdhulipa/Desktop/OneDrive - NREL/NREL-Github/EIA860PV-Wind/NARIS to Energy Visuals mapping/New Energy Visuals Data' 

Location_1 = Location +'/Nader/Energy-Visuals-SS-Data.xlsx'
Location_2 = Location +'/Nader/Energy-Visuals-Bus-Data.xlsx'
Location_3 = Location +'/Nader/Energy-Visuals-Gen-Data.xlsx'

Location_4 = Location +'/SS Mapping/EIA860-SolarPV-EV-SS-Mapping-FullList-New.xlsx'
Location_5 = Location +'/SS Mapping/EIA860-Wind-EV-SS-Mapping-FullList-New.xlsx'

df_ss = pd.read_excel(Location_1)
df_ss = df_ss.drop(df_ss.index[df_ss['# of Buses'] == 0])
df_bus = pd.read_excel(Location_2)
df_gen = pd.read_excel(Location_3)

# Solar
df_solar_mapping = pd.read_excel(Location_4)


df_solar_mapping['SS Bus Num'] = None
df_solar_mapping['SS Gen ID'] = None
df_solar_mapping['SS Gen Cap'] = None
for idx in df_solar_mapping.index:

    bus_idxs = df_bus.index[df_bus['Sub Num'] == df_solar_mapping['SS Number'].loc[idx]]
    bus_nums =[]
    gen_ids = []
    gen_caps = []        
    for idx_new in bus_idxs:
        bus_num = df_bus['Number'].loc[idx_new]
        temp_idx = df_gen.index[df_gen['Number of Bus'] == bus_num]
        if (len(temp_idx)!=0):
            bus_nums.append(bus_num)
            gen_ids.append(df_gen['ID'].loc[temp_idx[0]])
            gen_caps.append(df_gen['Max MW'].loc[temp_idx[0]])
        
    df_solar_mapping.at[idx,'SS Bus Num'] = bus_nums 
    df_solar_mapping.at[idx,'SS Gen ID'] = gen_ids   
    df_solar_mapping.at[idx,'SS Gen Cap'] = gen_caps      

# num_of_ss = []   

# for idx in df_solar_mapping.index:
#     idx = df_ss.index[df_ss['Sub Num'] == df_solar_mapping['SS Number'].loc[idx]][0]
#     num_of_ss.append(df_ss['# of Buses'].loc[idx])
    

# df_solar_mapping['Number of Buses'] = num_of_ss



bpat_ba_idx_pv=df_solar_mapping.index[df_solar_mapping['Balancing Authority Code']=="WALC"]

bpa_pv_plants = df_solar_mapping.loc[bpat_ba_idx_pv]


bpa_pv_plants['Other Close SS'] = None
bpa_pv_plants['Other Close SS Distance'] = None
bpa_pv_plants['Other Bus Numbers'] = None
bpa_pv_plants['Other Gen IDs'] = None
bpa_pv_plants['Other Gen Caps'] = None
for idx in bpa_pv_plants.index:
    
    if (len(bpa_pv_plants['SS Bus Num'].loc[idx])==0):
   
    
        gen_point = Point(bpa_pv_plants['Latitude'].loc[idx],bpa_pv_plants['Longitude'].loc[idx])
        
       
        geometry = [Point(x,y) for x,y in zip(df_ss['Latitude'],df_ss['Longitude'])]
        
        gen_bus_distance =[]
        
        
        for bus_point in geometry:
            gen_bus_distance.append(distance.distance(gen_point,bus_point).miles)
            
        
        sort_idx = np.argsort(gen_bus_distance)
        distance_min = gen_bus_distance[sort_idx[0]] 
        
        bpa_pv_plants.at[idx,'Other Close SS'] = df_ss['Sub Num'].iloc[sort_idx[1:6]].values
        bpa_pv_plants.at[idx,'Other Close SS Distance'] = np.array(gen_bus_distance)[sort_idx[1:6]]
        
        bus_nums_all = []
        gen_ids_all = []
        gen_caps_all = []
        for ss_num in bpa_pv_plants['Other Close SS'].loc[idx]:
            bus_idxs = df_bus.index[df_bus['Sub Num'] == ss_num]
            bus_nums =[]
            gen_ids = [] 
            gen_caps = []
            for idx_new in bus_idxs:
                bus_num = df_bus['Number'].loc[idx_new]
                temp_idx = df_gen.index[df_gen['Number of Bus'] == bus_num]
                if (len(temp_idx)!=0):
                    bus_nums.append(bus_num)
                    gen_ids.append(df_gen['ID'].loc[temp_idx[0]])
                    gen_caps.append(df_gen['Max MW'].loc[temp_idx[0]])
                    
            bus_nums_all.append(bus_nums)
            gen_ids_all.append(gen_ids)
            gen_caps_all.append(gen_caps)
                    
            
        bpa_pv_plants.at[idx,'Other Bus Numbers'] = bus_nums_all
        bpa_pv_plants.at[idx,'Other Gen IDs'] = gen_ids_all
        bpa_pv_plants.at[idx,'Other Gen Caps'] = gen_caps_all
        
        
    
# Wind    
    
df_wind_mapping = pd.read_excel(Location_5)

df_wind_mapping['SS Bus Num'] = None
df_wind_mapping['SS Gen ID'] = None
df_wind_mapping['SS Gen Cap'] = None

for idx in df_wind_mapping.index:

    bus_idxs = df_bus.index[df_bus['Sub Num'] == df_wind_mapping['SS Number'].loc[idx]]
    bus_nums =[]
    gen_ids = []
    gen_caps = []
    for idx_new in bus_idxs:
        bus_num = df_bus['Number'].loc[idx_new]
        temp_idx = df_gen.index[df_gen['Number of Bus'] == bus_num]
        if (len(temp_idx)!=0):
            bus_nums.append(bus_num)
            gen_ids.append(df_gen['ID'].loc[temp_idx[0]])
            gen_caps.append(df_gen['Max MW'].loc[temp_idx[0]])
        
    df_wind_mapping.at[idx,'SS Bus Num'] = bus_nums 
    df_wind_mapping.at[idx,'SS Gen ID'] = gen_ids
    df_wind_mapping.at[idx,'SS Gen Cap'] = gen_caps 
            
       
# num_of_ss = []   

# for idx in df_wind_mapping.index:
#     idx = df_ss.index[df_ss['Sub Num'] == df_wind_mapping['SS Number'].loc[idx]][0]
#     num_of_ss.append(df_ss['# of Buses'].loc[idx])
    

# df_wind_mapping['Number of Buses'] = num_of_ss

bpat_ba_idx_wind=df_wind_mapping.index[df_wind_mapping['Balancing Authority Code']=="WALC"]

bpa_wind_plants = df_wind_mapping.loc[bpat_ba_idx_wind]


bpa_wind_plants['Other Close SS'] = None
bpa_wind_plants['Other Close SS Distance'] = None
bpa_wind_plants['Other Bus Numbers'] = None
bpa_wind_plants['Other Gen IDs'] = None
bpa_wind_plants['Other Gen Caps'] = None

for idx in bpa_wind_plants.index:
    
     if (len(bpa_wind_plants['SS Bus Num'].loc[idx])==0):
    
    
        gen_point = Point(bpa_wind_plants['Latitude'].loc[idx],bpa_wind_plants['Longitude'].loc[idx])
        
       
        geometry = [Point(x,y) for x,y in zip(df_ss['Latitude'],df_ss['Longitude'])]
        
        gen_bus_distance =[]
        
        
        for bus_point in geometry:
            gen_bus_distance.append(distance.distance(gen_point,bus_point).miles)
            
        
        sort_idx = np.argsort(gen_bus_distance)
        distance_min = gen_bus_distance[sort_idx[0]] 
        
        bpa_wind_plants.at[idx,'Other Close SS'] = df_ss['Sub Num'].iloc[sort_idx[1:6]].values
        bpa_wind_plants.at[idx,'Other Close SS Distance'] = np.array(gen_bus_distance)[sort_idx[1:6]]
        
        bus_nums_all = []
        gen_ids_all = []
        gen_caps_all = []
        for ss_num in bpa_wind_plants['Other Close SS'].loc[idx]:
            bus_idxs = df_bus.index[df_bus['Sub Num'] == ss_num]
            bus_nums =[]
            gen_ids = []
            gen_caps = []
            for idx_new in bus_idxs:
                bus_num = df_bus['Number'].loc[idx_new]
                temp_idx = df_gen.index[df_gen['Number of Bus'] == bus_num]
                if (len(temp_idx)!=0):
                    bus_nums.append(bus_num)
                    gen_ids.append(df_gen['ID'].loc[temp_idx[0]])
                    gen_caps.append(df_gen['Max MW'].loc[temp_idx[0]])
                    
            bus_nums_all.append(bus_nums)
            gen_ids_all.append(gen_ids)
            gen_caps_all.append(gen_caps)      
            
        bpa_wind_plants.at[idx,'Other Bus Numbers'] = bus_nums_all
        bpa_wind_plants.at[idx,'Other Gen IDs'] = gen_ids_all
        bpa_wind_plants.at[idx,'Other Gen Caps'] = gen_caps_all

bpa_pv_plants.to_excel(Location +'/SS Mapping/BA-Level Mapping/Surya/WALC-EIA-860SolarPlant-SS-Mapping.xlsx', index=False)       
bpa_wind_plants.to_excel(Location +'/SS Mapping/BA-Level Mapping/Surya/WALC-EIA-860WindPlant-SS-Mapping.xlsx', index=False) 