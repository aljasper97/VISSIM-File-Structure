# -*- coding: utf-8 -*-
"""
VISSIM file packaging script
Created by: Amy Tibbetts 3/1/2023
"""

import os
import shutil

# Inputs
source_dir = r'\\nvafiler01\Project03\AT_NVA\Data\NVA_Vissim\Project_XXX'
PM = False
AM = False
my_dict_AM = {
           '3-2050 No-Build' : 'v22',
           '5-2050 Build Alt 1' : 'v25',
           '6-2050 Build Alt 2' : 'v22'
           }
my_dict_PM = {
           '3-2050 No-Build' : 'v27',
           '5-2050 Build Alt 1' : 'v18',
           '6-2050 Build Alt 2' : 'v10'
           }
l_filetypes = ['.inpx','.rbc'] # options: '.inpx' '.rbc' '.err'

# Processing starts here
if AM == True:
    
    for folder, version in my_dict_AM.items():
    
        file_path_am =  source_dir + '\\' + folder + '\\AM\\02-Outputs\\' + version
        to_dir_am = rf'{source_dir}\_Submittals\{folder}\AM'

        if os.path.exists(rf'{source_dir}\_Submittals\{folder}\AM'):
            shutil.rmtree(rf'{source_dir}\_Submittals\{folder}\AM') 

        for file in os.listdir(file_path_am):
            os.makedirs(to_dir_am, exist_ok=True) 
            
            for filetype in l_filetypes:
                if (file.endswith('inpx')) & (filetype == '.inpx'):
                    shutil.copy(os.path.join(file_path_am,file), os.path.join(to_dir_am, file))
                    oldName = os.path.join(to_dir_am,file)
                    newName = os.path.join(to_dir_am,file.replace(f"_{version}", ""))
                    os.rename(oldName, newName)
                    print(f'INFO - Copied {file} and renamed to', file.replace(f"_{version}", ""))
                
                if (file.endswith(".err")) & (filetype == '.err'):
                    shutil.copy(os.path.join(file_path_am,file), os.path.join(to_dir_am, file))
                    oldName = os.path.join(to_dir_am,file)
                    newName = os.path.join(to_dir_am,file.replace(f"_{version}", ""))
                    os.rename(oldName, newName)
                    print(f'INFO - Copied {file} and renamed to', file.replace(f"_{version}", ""))
                    
                if (file.endswith(".rbc")) & (filetype == '.rbc'):
                    shutil.copy(os.path.join(file_path_am,file), os.path.join(to_dir_am, file))
                    print('INFO - Copied', file)
                
if PM == True:
    
    for folder, version in my_dict_PM.items():

        file_path_pm =  source_dir + '\\' + folder + '\\PM\\02-Outputs\\' + version
        to_dir_pm = rf'{source_dir}\_Submittals\{folder}\PM'

        if os.path.exists(rf'{source_dir}\_Submittals\{folder}\PM'):
            shutil.rmtree(rf'{source_dir}\_Submittals\{folder}\PM') 

        for file in os.listdir(file_path_pm):
            os.makedirs(to_dir_pm, exist_ok=True) 
            
            for filetype in l_filetypes:
                if (file.endswith('inpx')) & (filetype == '.inpx'):
                    shutil.copy(os.path.join(file_path_pm,file), os.path.join(to_dir_pm, file))
                    oldName = os.path.join(to_dir_pm,file)
                    newName = os.path.join(to_dir_pm,file.replace(f"_{version}", ""))
                    os.rename(oldName, newName)
                    print(f'INFO - Copied {file} and renamed to', file.replace(f"_{version}", ""))
  
                if (file.endswith(".err")) & (filetype == '.err'):
                    shutil.copy(os.path.join(file_path_pm,file), os.path.join(to_dir_pm, file))
                    oldName = os.path.join(to_dir_pm,file)
                    newName = os.path.join(to_dir_pm,file.replace(f"_{version}", ""))
                    os.rename(oldName, newName)
                    print(f'INFO - Copied {file} and renamed to', file.replace(f"_{version}", ""))
                    
                if (file.endswith(".rbc")) & (filetype == '.rbc'):
                    shutil.copy(os.path.join(file_path_pm,file), os.path.join(to_dir_pm, file))
                    print('INFO - Copied', file)
            
