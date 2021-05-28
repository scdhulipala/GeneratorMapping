"""
Created on Feb 2021
@author: Quan Nguyen
Project: NAERM
Task: Automate the process of mapping wind and solar between EIA data and planning case
"""

import os, sys
import pandas
import shutil
import scipy, math, timeit
import win32com.client

import os
import re

from fuzzywuzzy import fuzz






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


  





# =============================================================================
# INPUT
# =============================================================================
Dir_Path                = os.path.dirname(os.path.abspath(__file__))

# NREL fist try mapping
NREL_1st_try_map_file   = r'\PNM-EIA-860SolarPlant-SS-Mapping.xlsx'

# PCM data
PCM_file                = r'C:\_QN\Projects\xxEnded\ABB_GRIDVIEW_TCF\Code\BA_Wind_Solar_summary\2030_Summary_by_BA.xlsx'
BA_name                 = 'PNM'

# Power flow case
PF_case_path            = r'C:\_QN\Projects\NAERM\SA7_BES2\Load_Forecast\WECC\WECC_PF_case'
PF_file                 = r'\WECC Apr 2019_20HS3ap-PWver21.pwb'

# Mapping parameters
Pmax_per_tol            = 10 # %





# =============================================================================
# MAIN
# =============================================================================
""" Read the maping data """
NREL_1st_try_df             = pandas.read_excel(Dir_Path + NREL_1st_try_map_file, sheet_name='Sheet1')

# EAI data
EIA_plant_names             = NREL_1st_try_df['Plant Name'].to_list()
EIA_plant_ID                = NREL_1st_try_df['Plant ID'].to_list()
EIA_plant_Pmax              = NREL_1st_try_df['Nameplate Capacity (MW)'].to_list()

# Exact macth in terms of location
exact_mapping_gen_bus_num   = NREL_1st_try_df['SS Bus Num'].to_list()
exact_mapping_gen_ID        = NREL_1st_try_df['SS Gen ID'].to_list()
exact_mapping_gen_Pmax      = NREL_1st_try_df['SS Gen Cap'].to_list()

# Find nearby matches
nearby_mapping_gen_bus_num  = NREL_1st_try_df['Other Bus Numbers'].to_list()
nearby_mapping_gen_ID       = NREL_1st_try_df['Other Gen IDs'].to_list()
nearby_mapping_gen_Pmax     = NREL_1st_try_df['Other Gen Caps'].to_list()


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












