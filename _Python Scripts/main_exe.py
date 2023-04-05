import yaml
import os, shutil, time, sys
import tkinter
from tkinter import filedialog, simpledialog
from vissim_simulation_classes_exe import Vissim_Network
from travel_time_plots import *
from organize_att_files import *
from run_excel_macro import *
from transit_metrics import *

root = tkinter.Tk()
root.title('Vissim Run Tool')
root.withdraw()

vissim_path = tkinter.filedialog.askopenfilename(title='Select Vissim File to Run - Please Follow the Naming Convention Project_Scenario_Period_Version.inpx')
if vissim_path == '':
    print('Vissim file not provided')
    sys.exit()

save_path_start = tkinter.filedialog.askdirectory(title='Select Output Folder')
if save_path_start == '':
    print('Output folder not provided')
    sys.exit()

# Read inputs from YAML
with open('vissim_inputs_exe.yaml') as f:
    conf = yaml.load(f, Loader=yaml.Loader)

# Run details of which functions to complete
run_vissim = conf['RUN_DETAILS']['run_vissim']
copy_moe_sheet = conf['RUN_DETAILS']['copy_moe_sheet']
transit_veh_records = conf['RUN_DETAILS']['transit_veh_records']
travel_time_plots = conf['RUN_DETAILS']['travel_time_plots']
clean_output_files = conf['RUN_DETAILS']['clean_output_files']
run_excel_macros = conf['RUN_DETAILS']['run_excel_macros']

# Paths - assuming standard file structure
open_path_start = conf['PROJECT_DETAILS']['vissim_path']
moe_spreadsheet_path = r'{}\_MOE Spreadsheet'.format(open_path_start)
layout_path = r'{}\_Layout Files'.format(open_path_start)
transit_data_path = r'{}\_Transit'.format(open_path_start)

# Initialize Log File
log_fn = 'calibration_logfile.txt'
log_file = open(os.path.join(save_path_start,log_fn),'a') 

# Derive Base File and Scenario Naming
fn = os.path.basename(vissim_path)
fn_end = fn[fn.rfind('_')+1:fn.rfind('.')]
if fn.find('AM') > 0:
    period = 'AM'
    node_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['node_collection_setup_AM']
    link_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['link_collection_setup_AM']
elif fn.find('PM') > 0:
    period = 'PM'
    node_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['node_collection_setup_PM']
    link_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['link_collection_setup_PM']
elif fn.find('SAT') > 0:
    period = 'SAT'
    node_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['node_collection_setup_SAT']
    link_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['link_collection_setup_SAT']


# Define the save path directory and create if needed
save_path = os.path.join(save_path_start,fn_end)
open_path = os.path.dirname(os.path.realpath(vissim_path))

if not os.path.exists(save_path):
    os.makedirs(save_path)
else:
    overwrite_file = tkinter.simpledialog.askstring(title="Alert!", prompt=f"Output folder {fn_end} already exists - do you want to overwrite it?  (Y/N):  ")
    if overwrite_file == "Y":  
        try:
            shutil.rmtree(save_path)
        except Exception as e:
            print(f'Error clearing path: {e}')
        if not os.path.exists(save_path):
            os.makedirs(save_path)
    else: # End script if you don't want to overwrite the folder.
        sys.exit()

if run_vissim is True:
    # Iterative process to archive and save-over all Model files
    print ("Copying over RBC and V3D Files")
        
    # Copy RBC files and Vehicle files to Output path/Running Path
    for f_ in os.listdir(open_path):
        if f_.endswith('.rbc') or f_.endswith('.v3d'):
            shutil.copyfile(src=open_path+'\{}'.format(f_),dst=save_path+'\{}'.format(f_))

    # Iterative process to run model(s)
    start_time = time.time()

    print ("Starting: {} at {}".format(fn, time.asctime(time.localtime(start_time))))
    log_file.write("Starting: {} at {}".format(fn, time.asctime(time.localtime(start_time))))
    log_file.write("\n")

    # Define the save path directory and create if needed
    if not os.path.exists(save_path):
        os.makedirs(save_path)

    # Open Vissim
    print ("Opening Vissim")
    network = Vissim_Network(network_name=fn, layout_name=conf['LAYX']['layout_fn'], \
        network_path=open_path, save_path=save_path, layout_path=layout_path, \
        version=conf['SIMULATION_PARAMETERS']['vissim_version'])
    network.save_network_files(new_save_path=save_path)
    
    # Setup Vissim Settings
    print ("Applying Vissim Settings")
    network.set_evaluation_output_directory(save_path=save_path)
    network.activate_quickmode()
    network.set_number_of_runs(no_runs=conf['SIMULATION_PARAMETERS']['no_runs'])
    network.set_run_time(run_time=conf['SIMULATION_PARAMETERS']['sim_run_time'])
    network.set_random_seed(random_seed = conf['SIMULATION_PARAMETERS']['random_seed_start'], 
                            increment = conf['SIMULATION_PARAMETERS']['random_seed_increment'])
    keep_prev_results = 2  # Keep Only of current (multiple) simulation
    network.set_evaluation(keep_prev_results = keep_prev_results,
            veh_class_recording = conf['VISSIM_EVALUTATION_PARAMETERS']['veh_class_recording'],
            data_collection_active = conf['VISSIM_EVALUTATION_PARAMETERS']['data_collection_active'], 
            data_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['data_collection_setup'],
            node_collection_active = conf['VISSIM_EVALUTATION_PARAMETERS']['node_collection_active'], 
            node_collection_setup = node_collection_setup,
            travel_time_collection_active = conf['VISSIM_EVALUTATION_PARAMETERS']['travel_time_collection_active'], 
            travel_time_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['travel_time_collection_setup'],
            veh_net_performance_active = conf['VISSIM_EVALUTATION_PARAMETERS']['veh_net_performance_active'],
            veh_net_performance_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['veh_net_performance_setup'],
            link_collection_active = conf['VISSIM_EVALUTATION_PARAMETERS']['link_collection_active'],
            link_collection_setup = link_collection_setup,
            queue_collection_active = conf['VISSIM_EVALUTATION_PARAMETERS']['queue_collection_active'],
            queue_collection_setup = conf['VISSIM_EVALUTATION_PARAMETERS']['queue_collection_setup'])
    
    # Run Vissim model
    print ("Running Vissim Model")
    network.run_complete_simulation()
    network.save_network_files(new_save_path=save_path)
    network.update_layout(layout_name=conf['LAYX']['layout_fn_post'])
    network.close()
    print ("Vissim Model Runs Complete")

    time.sleep(5)  # Wait loop for Vissim to shut down

# Clean .ATT Files (only keep the last)
if clean_output_files is True:
    print ("Clean up folder to remove unnecessary .ATT files")
    try: Clean_Output_Files(results_path=save_path, no_seeds=conf['SIMULATION_PARAMETERS']['no_runs'])
    except Exception as e:
        print("     *Error Cleaning .ATT Output Files")
        print(f"    Exception: {e}")

# Copy empty MOE spreadsheet to save path
if copy_moe_sheet is True:
    try:
        print ("Copying MOE Spreadsheet(s) into Path")
        l_moe_sheets = glob.glob(os.path.join(moe_spreadsheet_path,'*template*.xlsm'))  # Will copy all macro-enabled files with "template" in the name
        
        for moe_sheet in l_moe_sheets:
            moe_sheet_fn = os.path.basename(moe_sheet)
            moe_sheet_fn_new = os.path.basename(vissim_path).replace('.inpx','.xlsm')
            # Check if sheet exists in folder and if it does, rename it before copying
            if os.path.exists(save_path+'\{}'.format(moe_sheet_fn_new)):
                if not os.path.exists(save_path+'\!Archive'):
                    os.makedirs(save_path+'\!Archive')
                shutil.copyfile(src=save_path+'\{}'.format(moe_sheet_fn_new),dst=save_path+'\!Archive'+'\{}_{}'.format(time.strftime("%Y-%m-%d_%H-%M"),moe_sheet_fn_new))
            # Copy New File
            shutil.copyfile(src=moe_spreadsheet_path+'\{}'.format(moe_sheet_fn),dst=save_path+'\{}'.format(moe_sheet_fn_new))
    except Exception as e:
        print("     *Error Copying Out MOE Sheet")
        print(f"    Exception: {e}")

# Process Travel Time Data
if travel_time_plots is True:
    print ("Process Travel Time Data - Produce Travel Time Plots")
    try: Travel_Time_Plots(conf['TRAVEL_TIME_PLOTS']['TT_ylim'], results_path=save_path)
    except Exception as e:
        print("     *Error Reporting Travel Time Plots")
        print(f"    Exception: {e}")

# Process Transit Data
if transit_veh_records is True:
    print ("Process Transit Data - Produce Summary Spreadsheet")
    try: Transit_Data_Processing(version=fn_end, save_path=save_path, file_name_start=fn, 
                                    transit_data_path=transit_data_path, project_name = conf['PROJECT_DETAILS']['project_name'])
    except Exception as e:
        print("     *Error Reporting Transit Data")
        print(f"    Exception: {e}")

# Run MOE Sheet Macros
if run_excel_macros is True:
    print ("Run MOE Sheet Macros to Process Results")
    try: 
        l_moe_sheets = glob.glob(os.path.join(save_path,'*.xlsm'))
        for moe_sheet in l_moe_sheets:
            moe_sheet_fn = os.path.basename(moe_sheet)
            # Execute Macro to Clear all results, Toggle between AM and PM, then Process new .ATT results
            run_excel_macro(open_path=save_path, excel_name=moe_sheet_fn, excel_scenario=excel_scenario,
                            mod1="Main", macro1="Change_Links_Python",
                            mod2="Main", macro2="Clear_Results_Refresh",
                            mod3="Main", macro3="Import_Results_Auto")
    except Exception as e:
        print("     *Error Running MOE Sheet Macro")
        print(f"    Exception: {e}")

total_time = (time.time()-start_time)/60
print ("Finished - {} min".format(round(total_time,1)))
print ("")
log_file.write("Finished - {} min".format(round(total_time,1)))
log_file.write("\n")
log_file.write("\n")

log_file.close()

root.mainloop()

# User input to not close executable until acknowledged
close_decision = bool(input("Press enter to close "))