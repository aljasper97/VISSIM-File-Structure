# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import os
import shutil

# Inputs
source_dir = r'\\nvafiler01\Project03\AT_NVA\Data\NVA_Vissim\Project_XXX'
PM = True
AM = False
my_dict = {'1-Existing' : 'v0',
           '2-20XX No-Build' : 'v0',
           '3-20XX Build' : 'v0',
           }
l_filetypes = ['.inpx','.rbc'] # options: '.inpx' '.rbc' '.err'

# Processing starts here
for folder, version in my_dict.items():
    
    file_path_am =  source_dir + '\\' + folder + '\\AM\\02-Outputs\\' + version
    file_path_pm =  source_dir + '\\' + folder + '\\PM\\02-Outputs\\' + version
    to_dir_am = rf'{source_dir}\_Submittals\{folder}\AM'
    to_dir_pm = rf'{source_dir}\_Submittals\{folder}\PM'
    if os.path.exists(rf'{source_dir}\_Submittals\{folder}'):
        shutil.rmtree(rf'{source_dir}\_Submittals\{folder}') 
    
    if AM == True:
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
            
