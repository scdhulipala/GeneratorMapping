#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri May 28 13:48:00 2021

@author: sdhulipa
"""

###############################################################################
# Surya and Quan
# NREL and PNNL
# May 2021
# Script to map from generator lat/long to a set of lat/longs of substations by
# mapping a generator to the closest substation.
###############################################################################
# Importing Required Packages
###############################################################################
import pandas as pd
from geopy import distance, Point
import numpy as np

# Quan (PNNL)
import os, sys
import shutil
import scipy, math, timeit
import win32com.client

import re

from fuzzywuzzy import fuzz

# Flag to use PowerWorld Service
pw_flag = False

# Quan (PNNL ) Functions
def Remove_space(a_string):
    idx     = 0
    idx_del = []
    while idx < len(a_string)-1:
        if a_string[idx] == '[':
            for idx_next in range(idx, len(a_string)):
                if a_string[idx_next] == ']':
                    break
            for idx_mid in range(idx, idx_next+1):
                if a_string[idx_mid] == ' ':
                    idx_del.append(idx_mid)
            idx = idx_next+1
        else:
            idx += 1
    count = 0
    for a_del_index in idx_del:
        a_string = a_string[ : a_del_index-count] + a_string[a_del_index+1-count : ]
        count    += 1
    return a_string



def Convert_String_to_List(a_string, a_type):
    if (a_string[1]) != '[': # a single list
        if a_type == 'Bus Number':
            ans = map(int, a_string[1 : -1].split(','))
        elif a_type == 'Capacity':
            ans = map(float, a_string[1 : -1].split(','))
        elif a_type == 'ID':
            ans = map(str, a_string[1 : -1].split(', '))
            ans = [x[1 : -1] for x in ans]
        ans = [ans]
        
    else: # list of lists
        a_string    = Remove_space(a_string)
        a_list      = map(str, a_string[1 : -1].split(', '))
        ans         = []
        for x in a_list:
            if x == '[]':
                ans.append([])
            else:
                if a_type == 'Bus Number':
                    ans.append(map(int, x[1 : -1].split(',')))
                elif a_type == 'Capacity':
                    ans.append(map(float, x[1 : -1].split(',')))
                elif a_type == 'ID':
                    ans.append(map(str, x[1 : -1].split(',')))
        if a_type == 'ID':
            for list1 in ans:
                for idx in range(len(list1)):
                    list1[idx] = list1[idx][1 : -1] # need to debug to see this
    return ans





def match_plant_name(an_EIA_plant, a_PCM_plant_name):
    EIA_sub_string_list     = an_EIA_plant.replace('_',' ').split(' ')
    PCM_sub_string_list     = a_PCM_plant_name.replace('_',' ').split(' ')
    common_sub_string_list  = list(set(EIA_sub_string_list).intersection(PCM_sub_string_list))
    if len(common_sub_string_list) >= 2:
        print an_EIA_plant + ' in EIA is mapped to ' + a_PCM_plant_name + ' in PCM' 
        return True
    else:
        return False
    


def fuzzy_match_plant_name(an_EIA_plant, a_PCM_plant_name):
    Ratio = fuzz.ratio(an_EIA_plant.lower(),a_PCM_plant_name.lower())
    
    if Ratio > 60:   
        print '\n'+ an_EIA_plant + ' in EIA is FUZZY mapped to ' + a_PCM_plant_name + ' in PCM' 
        print 'String Match Ratio: ' + str(Ratio)
        return True
    else:
        return False
###############################################################################

# =============================================================================
# INPUT - Quan (PNNL)
# =============================================================================
# PCM data
PCM_file                = '2030_Summary_by_BA.xlsx'
BA_name                 = 'PNM'

# Power flow case
PF_file                 = r'\WECC Apr 2019_20HS3ap-PWver21.pwb'

# Mapping parameters
Pmax_per_tol            = 10 # %


Location_1 = 'EIA860-2019ER-April2020-SolarPV-List-Updated.xlsx' 
Location_2 = 'Energy-Visuals-SS-Data.xlsx'
Location_3 = Location +'Energy-Visuals-Bus-Data.xlsx'
Location_4 = Location +'Energy-Visuals-Gen-Data.xlsx'

# Reading the Excel Files

df_pv_plants = pd.read_excel(Location_1)
df_pv_plants = df_pv_plants.drop_duplicates(subset = ['Plant ID', 'Plant State'], keep = 'first').reset_index(drop = True)

df_ss_data = pd.read_excel(Location_2)
df_ss_data  = df_ss_data.loc[df_ss_data['# of Buses']!=0]

df_bus = pd.read_excel(Location_3)

df_gen = pd.read_excel(Location_4)

# Adding Necessary Columns for Mapping
df_pv_plants['SS Name'] = None
df_pv_plants['SS Number'] = None
df_pv_plants['Mapped EIA BA'] = None
df_pv_plants['Min Distance'] = None

df_pv_plants['SS Bus Num'] = None
df_pv_plants['SS Gen ID'] = None
df_pv_plants['SS Gen Cap'] = None

df_pv_plants['Other Close SS'] = None
df_pv_plants['Other Close SS Distance'] = None
df_pv_plants['Other Bus Numbers'] = None
df_pv_plants['Other Gen IDs'] = None
df_pv_plants['Other Gen Caps'] = None

for idx in df_pv_plants.index:
    
    gen_point = Point(df_pv_plants['Latitude'].loc[idx],df_pv_plants['Longitude'].loc[idx])
    geometry = [Point(x,y) for x,y in zip(df_ss_data['Latitude'],df_ss_data['Longitude'])]
    
    gen_bus_distance =[]
    
    for bus_point in geometry:
        gen_bus_distance.append(distance.distance(gen_point,bus_point).miles)
        
    
    sort_idx = np.argsort(gen_bus_distance)
    distance_min = gen_bus_distance[sort_idx[0]] 
   
    ss_idxs = df_ss_data.index[sort_idx[0:5]]
    df_pv_plants.at[idx,'SS Name'] = df_ss_data['Sub Name'].loc[ss_idxs[0]]
    df_pv_plants.at[idx,'SS Number'] = df_ss_data['Sub Num'].loc[ss_idxs[0]]
    df_pv_plants.at[idx, 'Mapped EIA BA'] =df_ss_data['Area Name'].loc[ss_idxs[0]]
    df_pv_plants.at[idx, 'Min Distance'] = distance_min
    
    bus_idxs = df_bus.index[df_bus['Sub Num'] == df_pv_plants['SS Number'].loc[idx]]
    
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
            
    df_pv_plants.at[idx,'SS Bus Num'] = bus_nums 
    df_pv_plants.at[idx,'SS Gen ID'] = gen_ids   
    df_pv_plants.at[idx,'SS Gen Cap'] = gen_caps
    
    df_pv_plants.at[idx,'Other Close SS'] = df_ss['Sub Num'].loc[ss_idxs[1:5]].values
    df_pv_plants.at[idx,'Other Close SS Distance'] = np.array(gen_bus_distance)[sort_idx[1:5]]
    
    bus_nums_all = []
    gen_ids_all = []
    gen_caps_all = []
    for ss_num in df_pv_plants['Other Close SS'].loc[idx]:
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
                
        
    df_pv_plants.at[idx,'Other Bus Numbers'] = bus_nums_all
    df_pv_plants.at[idx,'Other Gen IDs'] = gen_ids_all
    df_pv_plants.at[idx,'Other Gen Caps'] = gen_caps_all
    
   
# =============================================================================
# MAIN - Quan (PNNL)
# =============================================================================
""" Read the maping data """
# EIA Data
EIA_plant_names             = df_pv_plants['Plant Name'].to_list()
EIA_plant_ID                = df_pv_plants['Plant ID'].to_list()
EIA_plant_Pmax              = df_pv_plants['Nameplate Capacity (MW)'].to_list()

# Exact macth in terms of location
exact_mapping_gen_bus_num   = df_pv_plants['SS Bus Num'].to_list()
exact_mapping_gen_ID        = df_pv_plants['SS Gen ID'].to_list()
exact_mapping_gen_Pmax      = df_pv_plants['SS Gen Cap'].to_list()

# Find nearby matches
nearby_mapping_gen_bus_num  = df_pv_plants['Other Bus Numbers'].to_list()
nearby_mapping_gen_ID       = df_pv_plants['Other Gen IDs'].to_list()
nearby_mapping_gen_Pmax     = df_pv_plants['Other Gen Caps'].to_list()    

if (pw_flag):
    """ Read info from PW power planning case """
    simauto_obj                 = win32com.client.Dispatch("pwrworld.SimulatorAuto")
    simauto_output              = simauto_obj.OpenCase(PF_case_path + PF_file)
    if simauto_output[0] == '':
        print 'SUCCESSFUL OPEN THE PW PLANNING CASE'
    
    # Read the generator information from the power flow case
    object_type                 = 'Gen' #ObjectType
    param_list                  = ['BusNum', 'ID', 'MWMax'] #ParamList
    filter_name                 = '' # FilterName of all elements
    simauto_output              = simauto_obj.GetParametersMultipleElement(object_type, param_list, filter_name)
    if simauto_output[0] == '':
        print 'SUCCESSFUL READ GENERATOR DATA'
        
    PF_case_gen_bus_num         = []
    PF_case_gen_ID_num          = []
    PF_case_gen_gen_Pmax        = []
    for i in range(len(simauto_output[1])): #Number of key fields
        for j in range(len(simauto_output[1][i])): #Number of objects
            if i == 0:
                PF_case_gen_bus_num.append(int(simauto_output[1][i][j]))
            elif i == 1:
                PF_case_gen_ID_num.append(str(simauto_output[1][i][j]))
            elif i == 2:
                PF_case_gen_gen_Pmax.append(float(simauto_output[1][i][j]))
    PF_case_gen_bus_num         = scipy.array(PF_case_gen_bus_num)


""" Mapping directly between NREL data and the power flow case """
print '\n ------------- Start mapping ----------- \n'

print ' STEP 1: MAP USING THE NREL MAP AND PLANNING CASE USING EXACT LOCATION'
mapping_stt                 = {}
map_PF_gen_bus_num_list     = {}
map_PF_gen_ID_list          = {}
map_PF_gen_Pmax_list        = {}
min_Pmax_diff               = {}

# Map using the perfect location mapping
for EIA_index in range(len(EIA_plant_names)):
    print 'Processing EIA plant ' + str(EIA_index+1)
    an_EIA_plant            = EIA_plant_names[EIA_index]
    an_EIA_ID               = EIA_plant_ID[EIA_index]
    an_EIA_Pmax             = EIA_plant_Pmax[EIA_index]

    if len(exact_mapping_gen_bus_num[EIA_index]) != 2: # Empty bracket "[]"
        min_Pmax_diff[an_EIA_ID] = Pmax_per_tol
        PF_gen_bus_num_full     = Convert_String_to_List(exact_mapping_gen_bus_num[EIA_index], 'Bus Number')
        PF_gen_ID_full          = Convert_String_to_List(exact_mapping_gen_ID[EIA_index], 'ID')
        PF_gen_Pmax_full        = Convert_String_to_List(exact_mapping_gen_Pmax[EIA_index], 'Capacity')
        for PF_index in range(len(PF_gen_bus_num_full)):
            PF_gen_bus_num_list = PF_gen_bus_num_full[PF_index]
            PF_gen_ID_list      = PF_gen_ID_full[PF_index]
            PF_gen_Pmax_list    = PF_gen_Pmax_full[PF_index]
            
            # Approach 1: Consider each generator in the planning case SEPARATELY
            for gen_idx in range(len(PF_gen_bus_num_list)):
                PF_gen_bus_num  = PF_gen_bus_num_list[gen_idx]
                PF_gen_ID       = PF_gen_ID_list[gen_idx]
                PF_gen_Pmax     = PF_gen_Pmax_list[gen_idx]
                if abs(an_EIA_Pmax - PF_gen_Pmax)/float(an_EIA_Pmax) * 100 < min_Pmax_diff[an_EIA_ID]:
                    print '------ Find a good ONE-TO_ONE match with generators at same location for EIA plant ' + str(EIA_index+1) 
                    min_Pmax_diff[an_EIA_ID]  = abs(an_EIA_Pmax - PF_gen_Pmax)/float(an_EIA_Pmax) * 100
                    mapping_stt[an_EIA_ID] = 'Good one-to-one match with generators at same location'
                    map_PF_gen_bus_num_list[an_EIA_ID] = [PF_gen_bus_num]
                    map_PF_gen_ID_list[an_EIA_ID] = [PF_gen_ID ]
                    map_PF_gen_Pmax_list[an_EIA_ID] = [PF_gen_Pmax]
            
            # Approach 2: Consider all generators in the planning case TOGETHER
            if abs(an_EIA_Pmax - scipy.array(PF_gen_Pmax_list).sum(axis=0))/float(an_EIA_Pmax) * 100 < min_Pmax_diff[an_EIA_ID]:
                print '------ Find a good match with generators at same location for EIA plant ' + str(EIA_index+1) 
                min_Pmax_diff[an_EIA_ID]  = abs(an_EIA_Pmax - scipy.array(PF_gen_Pmax_list).sum(axis=0))/float(an_EIA_Pmax) * 100
                mapping_stt[an_EIA_ID] = 'Good match with generators at same location'
                map_PF_gen_bus_num_list[an_EIA_ID] = PF_gen_bus_num_list
                map_PF_gen_ID_list[an_EIA_ID] = PF_gen_ID_list
                map_PF_gen_Pmax_list[an_EIA_ID] = PF_gen_Pmax_list
        if min_Pmax_diff[an_EIA_ID] == Pmax_per_tol:
            print 'Can not find a good match with generators at same location for EIA plant ' + str(EIA_index+1) 
    else:
        print 'Can not find a good match with generators at same location for EIA plant ' + str(EIA_index+1) 



# Map using the nearby location mapping
print '\n \n'
print ' STEP 2: MAP USING THE NREL MAP AND PLANNING CASE USING NEARBY LOCATION'
for EIA_index in range(len(EIA_plant_names)):
    print 'Processing EIA plant ' + str(EIA_index+1)
    an_EIA_plant            = EIA_plant_names[EIA_index]
    an_EIA_ID               = EIA_plant_ID[EIA_index]
    an_EIA_Pmax             = EIA_plant_Pmax[EIA_index]
    
    if str(nearby_mapping_gen_bus_num[EIA_index]) != 'nan':
        if an_EIA_ID not in min_Pmax_diff.keys():
            min_Pmax_diff[an_EIA_ID] = Pmax_per_tol
        PF_gen_bus_num_full     = Convert_String_to_List(nearby_mapping_gen_bus_num[EIA_index], 'Bus Number')
        PF_gen_ID_full          = Convert_String_to_List(nearby_mapping_gen_ID[EIA_index], 'ID')
        PF_gen_Pmax_full        = Convert_String_to_List(nearby_mapping_gen_Pmax[EIA_index], 'Capacity')
        for PF_index in range(len(PF_gen_bus_num_full)):
            PF_gen_bus_num_list = PF_gen_bus_num_full[PF_index]
            PF_gen_ID_list      = PF_gen_ID_full[PF_index]
            PF_gen_Pmax_list    = PF_gen_Pmax_full[PF_index]
            
            # Approach 1: Consider each generator in the planning case SEPARATELY
            for gen_idx in range(len(PF_gen_bus_num_list)):
                PF_gen_bus_num  = PF_gen_bus_num_list[gen_idx]
                PF_gen_ID       = PF_gen_ID_list[gen_idx]
                PF_gen_Pmax     = PF_gen_Pmax_list[gen_idx]
                if abs(an_EIA_Pmax - PF_gen_Pmax)/float(an_EIA_Pmax) * 100 < min_Pmax_diff[an_EIA_ID]:
                    print '------ Find a good ONE-TO_ONE match with generators at same location for EIA plant ' + str(EIA_index+1)
                    min_Pmax_diff[an_EIA_ID] = abs(an_EIA_Pmax - PF_gen_Pmax)/float(an_EIA_Pmax) * 100
                    mapping_stt[an_EIA_ID] = 'Good one-to-one match with generators at same location'
                    map_PF_gen_bus_num_list[an_EIA_ID] = [PF_gen_bus_num]
                    map_PF_gen_ID_list[an_EIA_ID] = [PF_gen_ID ]
                    map_PF_gen_Pmax_list[an_EIA_ID] = [PF_gen_Pmax]
            
            # Approach 2: Consider all generators in the planning case TOGETHER
            if abs(an_EIA_Pmax - scipy.array(PF_gen_Pmax_list).sum(axis=0))/float(an_EIA_Pmax) * 100 < min_Pmax_diff[an_EIA_ID]:
                print '------ Find a good match with generators at nearby location for EIA plant ' + str(EIA_index+1) 
                min_Pmax_diff[an_EIA_ID]  = abs(an_EIA_Pmax - scipy.array(PF_gen_Pmax_list).sum(axis=0))/float(an_EIA_Pmax) * 100
                mapping_stt[an_EIA_ID] = 'Good match with generators at nearby location'
                map_PF_gen_bus_num_list[an_EIA_ID] = PF_gen_bus_num_list
                map_PF_gen_ID_list[an_EIA_ID] = PF_gen_ID_list
                map_PF_gen_Pmax_list[an_EIA_ID] = PF_gen_Pmax_list
        if min_Pmax_diff[an_EIA_ID] == Pmax_per_tol:
            print 'Can not find a good match with generators at nearby location for EIA plant ' + str(EIA_index+1) 
    else:
        print 'Can not find a good match with generators at nearby location for EIA plant ' + str(EIA_index + 1) 






# Mapping using the PCM model
print '\n \n'
print ' STEP 3: MAP USING THE NREL MAP, 2030 PCM, AND PLANNING CASE'
PCM_df              = pandas.read_excel(PCM_file, sheet_name = BA_name)
PCM_plant_busnum    = PCM_df['Bus Number'].to_list()
PCM_plant_name      = PCM_df['Project Name'].to_list()
PCM_plant_Pmax      = PCM_df['Capacity (MW)'].to_list()

for EIA_index in range(len(EIA_plant_names)):
    print '\nProcessing EIA plant ' + str(EIA_index + 1)
    an_EIA_plant    = EIA_plant_names[EIA_index]
    an_EIA_ID       = EIA_plant_ID[EIA_index]
    an_EIA_Pmax     = EIA_plant_Pmax[EIA_index]
    if an_EIA_ID not in min_Pmax_diff.keys():
        min_Pmax_diff[an_EIA_ID] = Pmax_per_tol
        
    for PCM_index in range(len(PCM_plant_busnum)):
        a_PCM_plant_name    = PCM_plant_name[PCM_index]
        a_PCM_plant_busnum  = PCM_plant_busnum[PCM_index]
#        a_PCM_plant_Pmax    = PCM_plant_Pmax[PCM_index]
        if (match_plant_name(an_EIA_plant, a_PCM_plant_name)) or (fuzzy_match_plant_name(an_EIA_plant, a_PCM_plant_name)):    
            idx_PF      = scipy.where(PF_case_gen_bus_num == a_PCM_plant_busnum)[0]
            if len(idx_PF) > 0:
                idx_PF      = idx_PF[0]
                PF_gen_bus_num_list = PF_case_gen_bus_num[idx_PF]
                PF_gen_ID_list      = PF_case_gen_ID_num[idx_PF]
                PF_gen_Pmax_list    = PF_case_gen_gen_Pmax[idx_PF]
                if abs(an_EIA_Pmax - PF_gen_Pmax_list)/float(an_EIA_Pmax) * 100 < min_Pmax_diff[an_EIA_ID]:
                    print ' ---- Find a good match using PCM data for EIA plant ' + str(EIA_index+1) 
                    min_Pmax_diff[an_EIA_ID]  = abs(an_EIA_Pmax - PF_gen_Pmax_list)/float(an_EIA_Pmax) * 100
                    mapping_stt[an_EIA_ID] = 'Good match using PCM data'
                    map_PF_gen_bus_num_list[an_EIA_ID] = PF_gen_bus_num_list
                    map_PF_gen_ID_list[an_EIA_ID] = PF_gen_ID_list
                    map_PF_gen_Pmax_list[an_EIA_ID] = PF_gen_Pmax_list
                
    if min_Pmax_diff[an_EIA_ID] == Pmax_per_tol:
        print 'Can not find a good match using PCM data for EIA plant ' + str(EIA_index+1) 

        
