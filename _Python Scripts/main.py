# -*- coding: utf-8 -*-
"""
Created on Mon May 17 13:07:44 2021

@author: Ilana.Sadholz
"""

from vissim_run_functions import *


# =============================================================================
# INPUT: Indicate versions to run for each scenario:
# =============================================================================
l_s01 = ['v0'] # 1-2024 No-Build AM
l_s02 = [] # 1-2024 No-Build PM
l_s03 = [] # 2-2024 Build Refined F1 AM
l_s04 = [] # 2-2024 Build Refined F1 PM
l_s05 = [] # 3-2024 Build Beautiful Blvd AM
l_s06 = [] # 3-2024 Build Beautiful Blvd PM
l_s07 = [] 
l_s08 = [] 
l_s09 = [] 
l_s10 = [] 
l_s11 = [] 
l_s12 = [] 
l_s13 = [] 
l_s14 = [] 
l_s15 = [] 
l_s16 = []
l_s17 = [] 
l_s18 = [] 
l_s19 = [] 
l_s20 = [] 

# =============================================================================
# INPUT: Adjust these True/False variables based on what needs to be run
# =============================================================================
define_run = 3 # 1 - run everything, 2 - process moe's

if define_run == 1:
    run_vissim = True
    copy_moe_sheet = True
    transit_veh_records = False
    travel_time_plots = True
    clean_output_files = True
    run_excel_macros = True
    
elif define_run == 2:  # Change in MOE Sheet Update in Outputs
    run_vissim = False
    copy_moe_sheet = True
    transit_veh_records = False
    travel_time_plots = False
    clean_output_files = False
    run_excel_macros = True

elif define_run == 3:  #Testing
    run_vissim = True
    copy_moe_sheet = True
    transit_veh_records = False
    travel_time_plots = True
    clean_output_files = True
    run_excel_macros = False

# =============================================================================
# PROJECT SPECIFIC INPUTS
# =============================================================================
 
# Simulation parameters =======================================================
no_runs = 2  # Number of Random Seeds
vissim_version = "Vissim.Vissim-64.1100"  # Using 11.00-14, Last updated 2019-11-19 (Most recent update at start of project)
sim_run_time = 2400  # Seeding, Peak Hour, Shoulder
random_seed_start = 100
random_seed_increment = 1

# Evaluation Parameters =======================================================
keep_prev_results = 2  # Keep Only of current (multiple) simulation
# TO DO: try and figure out selecting the classes

# PARAMETERS: 
veh_class_recording = False # 999:All Motorized Vehicles - Set Manually in Vissim - IGNORE This line

data_collection_active = False
data_collection_setup = [0,99999,99999] # from time, to time, interval

node_collection_active = True
node_collection_setup_AM = [1800,5400,3600] # from time, to time, interval
node_collection_setup_PM = [1800,5400,3600]

travel_time_collection_active = True
travel_time_collection_setup = [0,7200,900] # from time, to time, interval

veh_net_performance_active = True
veh_net_performance_setup = [0,7200,7200] 

link_collection_active = False
link_collection_setup_AM = [0,99999,99999] # from time, to time, interval
link_collection_setup_PM = [0,99999,99999]

queue_collection_active = False  # Queue Counters
queue_collection_setup = [0,99999,99999] # from time, to time, interval

# Define y limit for TT plots
TT_ylim = 600


#!!!autosave_after_simulation = 1  # AutoExportTypeFile - DO THIS MANUALLY in Vissim
#!!!show_sim_run_aggregates = 1 # Show Simulation Run Aggregates (ave, std) - DO THIS MANUALLY in Vissim
#!!!Set Transit settings manually in Vissim

# Layout file =================================================================
layout_fn ='Outputs_MasterLayout_NHTP'  # Loads as model is running
layout_fn_post = 'Outputs_MasterLayout_NHTP_Post'  # Loads after runs are complete


# -----------------------------------------------------------------------
# Create dictionary from scenarios/versions to be run
d_scenarios = {'l_s01': l_s01, 'l_s02': l_s02, 
               'l_s03': l_s03, 'l_s04': l_s04, 
               'l_s05': l_s05, 'l_s06': l_s06, 
               'l_s07': l_s07, 'l_s08': l_s08,
               'l_s09': l_s09, 'l_s10': l_s10,
               'l_s11': l_s11, 'l_s12': l_s12,
               'l_s13': l_s13, 'l_s14': l_s14,
               'l_s15': l_s15, 'l_s16': l_s16,
               'l_s17': l_s17, 'l_s18': l_s18,
               'l_s19': l_s19, 'l_s20': l_s20}

for key, value in d_scenarios.items(): # Looping through each scenario
    if len(value) > 0:    
        for fn_end in value: # Looping through each version of that scenario
            # Defines scenario-specific variables
            # p_scenario: scenario name in file path
            # peak_pd: peak period of the model, AM or PM
            # fn_start: Key word for naming of the MOE sheet (replaces the "empty")
                
            if key == 'l_s01':
                p_scenario = '1-2024 No-Build' # string of path name for this scenario
                fn_scenario = "2024NB" # string of file name for this scenario
                peak_pd = 'AM' # String of peak period for this scenario
                excel_scenario = "2024 No-Build AM"
                
            if key == 'l_s02':
                p_scenario = '1-2024 No-Build' # string of path name for this scenario
                fn_scenario = "2024NB" # string of file name for this scenario
                peak_pd = 'PM' # String of peak period for this scenario
                excel_scenario = "2024 No-Build PM"
                
            if key == 'l_s03':
                p_scenario = '2-2024 Build Refined F1' # string of path name for this scenario
                fn_scenario = "2024Build1" # string of file name for this scenario
                peak_pd = 'AM' # String of peak period for this scenario
                excel_scenario = "2024 Build 1 AM"
                
            if key == 'l_s04':
                p_scenario = '2-2024 Build Refined F1' # string of path name for this scenario
                fn_scenario = "2024Build1" # string of file name for this scenario
                peak_pd = 'PM' # String of peak period for this scenario
                excel_scenario = "2024 Build 1 PM"
                
            if key == 'l_s05':
                p_scenario = '3-2024 Build Beautiful Blvd' # string of path name for this scenario
                fn_scenario = "2024Build2" # string of file name for this scenario
                peak_pd = 'AM' # String of peak period for this scenario
                excel_scenario = "2024 Build 2 AM"
                
            if key == 'l_s06':
                p_scenario = '3-2024 Build Beautiful Blvd' # string of path name for this scenario
                fn_scenario = "2024Build2" # string of file name for this scenario
                peak_pd = 'PM' # String of peak period for this scenario
                excel_scenario = "2024 Build 2 PM"
                
            if key == 'l_s07':
                p_scenario = '' # string of path name for this scenario
                fn_scenario = "" # string of file name for this scenario
                peak_pd = '' # String of peak period for this scenario
                excel_scenario = ""
                
            if key == 'l_s08':
                p_scenario = '' # string of path name for this scenario
                fn_scenario = "" # string of file name for this scenario
                peak_pd = '' # String of peak period for this scenario
                excel_scenario = ""
                
            if key == 'l_s09':
                p_scenario = '' # string of path name for this scenario
                fn_scenario = "" # string of file name for this scenario
                peak_pd = '' # String of peak period for this scenario
                excel_scenario = ""
                
            if key == 'l_s10':
                p_scenario = '' # string of path name for this scenario
                fn_scenario = "" # string of file name for this scenario
                peak_pd = '' # String of peak period for this scenario
                excel_scenario = ""

            if key == 'l_s11':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s12':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s13':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s14':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s15':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s16':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s17':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s18':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s19':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                
            if key == 'l_s20':
                p_scenario = ""
                fn_scenario = ""
                peak_pd = ""
                excel_scenario = ""
                                
             # set node collection times corresponding to peak period
            if peak_pd == "AM":
                node_collection_setup = node_collection_setup_AM
                link_collection_setup = link_collection_setup_AM
                
            if peak_pd == "PM":
                node_collection_setup = node_collection_setup_PM
                link_collection_setup = link_collection_setup_PM
            
            # Creating paths and filenames from scenario inputs & Running model
            fn_start = 'NHTP_{}_{}_'.format(fn_scenario, peak_pd)
            
            # File paths and file names ===================================================
            open_path = r'C:\Users\britton.hammit\OneDrive - KH\Projects\0-Gobys&Tools\Python\NHTP_Vissim\{}\{}\04-Working_Models'.format(p_scenario, peak_pd)
            save_path_start = r'C:\Users\britton.hammit\OneDrive - KH\Projects\0-Gobys&Tools\Python\NHTP_Vissim\{}\{}\05-Outputs'.format(p_scenario, peak_pd)
            
            moe_spreadsheet_fn = 'NHTP_MOE_template.xlsm'
            moe_spreadsheet_path = r'"C:\Users\britton.hammit\OneDrive - KH\Projects\0-Gobys&Tools\Python\NHTP_Vissim\_MOE Spreadsheet"'
            layout_path = moe_spreadsheet_path
            #transit_data_path = r'K:\NVA_RDWY\110721000_Route1_Multimodal\Production\Task 2 - Mulitmodal Transportation Analysis\04-Vissim\08-Transit'
        
            # Call function to run Vissim for each specified scenario/version
            RunVissim(fn_scenario,
                      peak_pd,
                      excel_scenario,
                      fn_start,
                      fn_end,
                      run_vissim,
                      copy_moe_sheet,
                      transit_veh_records,
                      travel_time_plots, 
                      clean_output_files,
                      run_excel_macros,
                      no_runs,
                      vissim_version,
                      sim_run_time,
                      random_seed_start,
                      random_seed_increment,
                      keep_prev_results,
                      veh_class_recording,
                      data_collection_active,
                      data_collection_setup,
                      node_collection_active, 
                      node_collection_setup,
                      travel_time_collection_active,
                      travel_time_collection_setup,
                      veh_net_performance_active,
                      veh_net_performance_setup,
                      link_collection_active,
                      link_collection_setup,
                      queue_collection_active,
                      queue_collection_setup,
                      TT_ylim,
                      layout_fn,
                      layout_fn_post,
                      open_path,
                      save_path_start,
                      moe_spreadsheet_fn,
                      moe_spreadsheet_path,
                      layout_path)
        

