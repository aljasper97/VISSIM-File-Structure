# -*- coding: utf-8 -*-
"""
Kimley-Horn
@authors: Britton Hammit

Extracts transit records from FZP files
* Produces summary table of transit travel time by route
* Produces summary table of transit dwell time by stop

Instructions for setting up Vissim
* Evaluation / Direct Output / Vehicle record 
* Check to "Write to File" and set up time boundaries as desired
* Select "More"
* Filter by vehicle class "Bus"
* Select Attributes - doesn't really matter how many different attributes are selected
*    the ones that are necessary are: Stop, Line, Dwell Time
* The header variable in the df_data read-in line will need to be updated accordingly.
* Change evaluation path

Last Update: 9/2/2021, Britton Hammit
"""

import pandas as pd
import os, glob


def Transit_Data_Processing(version,save_path,file_name_start,transit_data_path,project_name="project_name"):
    
    # Save Results
    writer = pd.ExcelWriter(os.path.join(save_path,"{}_transit.xlsx".format(file_name_start).replace(".fzp","")), engine='xlsxwriter')
    
    l_file_names = glob.glob(os.path.join(save_path,'*.fzp*'))
    
    l_df_stops_results = []
    l_df_lines_results = []
    
    for input_file in l_file_names:
        seed_no = input_file.replace(save_path,"")
        seed_no = seed_no.replace(".fzp","")
        seed_no = seed_no.replace(file_name_start,"")
        seed_no = seed_no.replace("\_","")
        
        # Determine Header Row (Line after the last asterisk)
        i=0
        with open(input_file) as f:
            for line in f:
               if "$VEHICLE:SIMSEC" in line: # indicator of header row
                    break
               else:
                    i+=1
        # Open file based on identified header row i
        df_data = pd.read_csv(filepath_or_buffer=input_file,delimiter=";",header = i, dtype = {"PTSTOP":float,"PTLINE":float},encoding='windows-1252')
        del i
        
        # Read in Key Data for each stop and line
        
        fn_pt_lines = "{}_PT_Line_Name.csv".format(project_name)
        fn_pt_stops = "{}_PT_Stop_Name.csv".format(project_name)
        
        df_lines = pd.read_csv(filepath_or_buffer=os.path.join(transit_data_path,fn_pt_lines),delimiter=",",index_col=0,header=0,dtype={"PTLINE":float})
        df_stops = pd.read_csv(filepath_or_buffer=os.path.join(transit_data_path,fn_pt_stops),delimiter=",",index_col=0,header=0,dtype={"PTSTOP":float})
        
        # Identify the Actual Dwell Time for each vehicle at each stop and average
        df_stops_results = pd.read_csv(filepath_or_buffer=os.path.join(transit_data_path,fn_pt_stops),delimiter=",",index_col=0,header=0,dtype={"PTSTOP":float})
        
        for index, series in df_stops.iterrows():
            # index = PT stop
            df_temp = df_data[df_data['PTSTOP']==index]  # filter to data from this stop
            df_temp = df_temp.sort_values(by=["DWELLTM"],ascending=False)
            df_temp.drop_duplicates(subset="NO",keep="first",inplace=True)
            df_stops_results.at[index,"Average DWELLTM"]=df_temp['DWELLTM'].mean()
            df_stops_results.at[index,"Stdev DWELLTM"]=df_temp['DWELLTM'].std()
            df_stops_results.at[index,"95th DWELLTM"]=df_temp['DWELLTM'].quantile(q=0.95)
            df_stops_results.at[index,"Max DWELLTM"]=df_temp['DWELLTM'].max()
            df_stops_results.at[index,"5th DWELLTM"]=df_temp['DWELLTM'].quantile(q=0.05)
            df_stops_results.at[index,"Min DWELLTM"]=df_temp['DWELLTM'].min()
            df_stops_results.at[index,"Count DWELLTM"]=df_temp['DWELLTM'].count()
        
        # Identify the Average Travel Time for each Bus Line
        df_lines_results = pd.read_csv(filepath_or_buffer=os.path.join(transit_data_path,fn_pt_lines),delimiter=",",index_col=0,header=0,dtype={"PTLINE":float})
        
        for index, series in df_lines.iterrows():
            # index = PT stop
            df_temp = df_data[df_data['PTLINE']==index]  # filter to data from this line
            df_temp = df_temp.sort_values(by=["TMINNETTOT"],ascending=False)  # sort by total time in network
            df_temp.drop_duplicates(subset="NO",keep="first",inplace=True)
            df_lines_results.at[index,"Average TT"]=df_temp['TMINNETTOT'].mean()
            df_lines_results.at[index,"Stdev TT"]=df_temp['TMINNETTOT'].std()
            df_lines_results.at[index,"95th TT"]=df_temp['TMINNETTOT'].quantile(q=0.95)
            df_lines_results.at[index,"Max TT"]=df_temp['TMINNETTOT'].max()
            df_lines_results.at[index,"5th TT"]=df_temp['TMINNETTOT'].quantile(q=0.05)
            df_lines_results.at[index,"Min TT"]=df_temp['TMINNETTOT'].min()
            df_lines_results.at[index,"Count TT"]=df_temp['TMINNETTOT'].count()
            
        df_stops_results.to_excel(writer, sheet_name='{}_Stops_Dwell'.format(seed_no), index = True)
        df_lines_results.to_excel(writer, sheet_name='{}_Lines_TT'.format(seed_no), index = True)
    
        l_df_stops_results.append(df_stops_results)
        l_df_lines_results.append(df_lines_results)
        
    
    df_stops_results_all = pd.concat(l_df_stops_results).groupby(level=0).mean()
    df_stops_results_all.to_excel(writer, sheet_name='ALL_Stops_Dwell', index = True)
    df_lines_results_all = pd.concat(l_df_lines_results).groupby(level=0).mean()
    df_lines_results_all.to_excel(writer, sheet_name='ALL_Lines_TT', index = True)
    
    writer.save()


def main():
    version = "v23"
    peak_pd = "AM"
    #peak_pd = "04-PM"
    save_path = r'K:\NVA_Vissim\Route_1\Vissim\01-Existing\{}\02-Outputs\{}'.format(peak_pd, version)
    transit_data_path = r'K:\NVA_Vissim\Route_1\Vissim\08-Transit'
    file_name_start = 'Route1_Existing_AM_{}'.format(version)
    Transit_Data_Processing(version,save_path,file_name_start,transit_data_path,'Route1')

if __name__ == "__main__":
    main()
    

#test_stop_no = 1001178
#df_temp = df_data[df_data['PTSTOP']==test_stop_no]  # filter to data from this stop
#df_temp.drop_duplicates(subset="NO",keep="last",inplace=True)