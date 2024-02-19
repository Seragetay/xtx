
import pandas as pd
import os
import duckdb
import glob
import shutil

# Top code runs every monday evenining
# using getlogin() returning username
user_name = os.getlogin()

def find_last_n_matching_files(root_folder, file_ext, contains, n):
    files = []

    for filename in os.listdir(root_folder):
        if filename.endswith(file_ext) and contains in filename and "$" not in filename:
            file_path = os.path.join(root_folder, filename)
            create_time = os.path.getctime(file_path)
            files.append((file_path, create_time))

    # Sort files by creation time in descending order
    sorted_files = sorted(files, key=lambda x: x[1], reverse=True)

    # Get the last n files
    last_n_files = [file[0] for file in sorted_files[:n]]

    return last_n_files

member_attr_root = r"W:\STARS_2024\RxAnte\Reports\KPIs" #choose the path for the current year
latest_mem_att_ppo_file =  find_last_n_matching_files(member_attr_root, ".xlsx", "KPI_Deliver",1) #choose n for the number of latest files
print(f' Loading....{latest_mem_att_ppo_file}')


folder_path = fr"C:\Users\{user_name}\Blue Cross Blue Shield of Michigan\StarsTeam-EH - General\5. Star Data and Important Docs\Stars Analytics\Pharmacy\Rx_Ante\temp_data"

for file_path in latest_mem_att_ppo_file:
    shutil.copy(file_path, folder_path)

xlsx_files = glob.glob(os.path.join(folder_path, '*KPI_Deliver*.xlsx'))

dfs = []

for file in latest_mem_att_ppo_file:
    df = pd.read_excel(file, skiprows=range(0,3))
    df = df.dropna(axis= 1, how = 'all')
    df['File_Name'] = os.path.basename(file)
    dfs.append(df)

combined_df = pd.concat(dfs, ignore_index=True)

######### append the data to this duckdb ############

conn_string = fr"C:\Users\{user_name}\Blue Cross Blue Shield of Michigan\StarsTeam-EH - General\5. Star Data and Important Docs\Stars Analytics\Database\pharmacy_data.db"
conn = duckdb.connect(conn_string)

# BUILD SOME CODE TO PREVENT DUPLICATIONS
conn.sql('INSERT INTO rx_ante.KPI BY NAME SELECT  * FROM combined_df')

# conn.execute("SHOW ALL TABLES").df() 
conn.commit()


conn.close()

for i in xlsx_files:
    shutil.os.remove(i)
    print(f'{i} deleted')



################################## RxAnte Patient list ######################
# Runs every Wed evening

import os 
import pandas as pd
import duckdb
from glob import glob
import re
from tqdm import tqdm
import numpy as np

# Specify the path
path = r"W:\STARS_2024\RxAnte\Reports\Weekly Files" #choose the path for the current year


# Connect to duckdb database

conn_string = fr"C:\Users\{user_name}\Blue Cross Blue Shield of Michigan\StarsTeam-EH - General\5. Star Data and Important Docs\Stars Analytics\Database\pharmacy_data.db"
conn = duckdb.connect(conn_string)




# Get all xlsx files containing "RxAnte" but excluding specified dates
existing_files_query = "SELECT DISTINCT FileName FROM rx_ante.rxante_patient_list"
existing_files = conn.execute(existing_files_query).df()

new_files_to_load = [
    file
    for file in glob(os.path.join(path, "**", "*RxAnte_Patient_List*.xlsx"), recursive=True)
    if "20230208" not in file and "20230215" not in file and "20230222" not in file and "~" not in file
    and os.path.basename(file) not in existing_files["FileName"].tolist()
]


# If there are new files to load, proceed with loading them into the database
if new_files_to_load:
    # Create an empty DataFrame
    combined_df_new = pd.DataFrame()

    # Read each new file, combine into one DataFrame, and add filename as a column
    for file in new_files_to_load:
        df = pd.read_excel(file, sheet_name= "ALL_MAPD", header=[0,1])
        df.columns = df.columns.map('_'.join)
        df.columns = df.columns.str.replace("Back to Contents_", "")
        df = df[df['Contract ID'].notnull()]
        df["FileName"] = os.path.basename(file)
        pattern = r"(.{8})\.xlsx$"
        match = re.search(pattern, file ) 
        extracted_char = match.group(1)
        df['file_date'] = extracted_char
        df['file_date'] = pd.to_datetime(df['file_date'])
        combined_df_new = combined_df_new.append(df, ignore_index=True)
        print(f' Added....{extracted_char} file to combined_df_new')

    # Insert the new data into the duckdb database
    conn.execute("INSERT INTO rx_ante.rxante_patient_list SELECT * FROM combined_df_new")



conn.commit()

conn.close()

print('program ran successfully')











#################################################################################


import pandas as pd
import os
import re
import win32com.client as win32
import numpy as np
from datetime import datetime, timedelta, date
import shutil
import glob
import warnings
warnings.filterwarnings("ignore")
pd.set_option('display.max_columns', None)



def find_patient_list_files(root_folder):
    # Create a list to store the matching file paths
    matching_files = []
    
    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".xlsx") and "patient_list" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    
    return matching_files

# Specify the root folder where the search should start
root_folder = r'W:\STARS_2023\RxAnte\Reports\Weekly Files'

# Call the function to find the matching files
matching_files = find_patient_list_files(root_folder)



file_path = matching_files[-1]
print(f' Loading....{file_path}')
type(file_path)
pattern = r"(.{8})\.xlsx$"
match = re.search(pattern, file_path ) 
extracted_char = match.group(1)
print(f' Loading file date....{extracted_char}')

#workbook = pd.ExcelFile(r"W:\STARS_2023\RxAnte\Reports\Weekly Files\20230607\RxAnte_patient_list_20230607.xlsx")
#workbook.sheet_names


#test_file_path = r"W:\PDE\ BA451-14 CMS Audit ESRD_d111315.xlsx"
#test_df = pd.read_excel(test_file_path)

print(f"reading Patient list from {extracted_char}")
#%time
df = pd.read_excel(file_path, sheet_name= "ALL_MAPD", header=[0,1])


#df = pd.concat([chunk for chunk in tqdm(pd.read_excel(file_path, sheet_name= "ALL_MAPD", header=[0,1], chunksize=1000), desc='Loading data')])

df.columns = df.columns.map('_'.join)

df.columns = df.columns.str.replace("Back to Contents_", "")

df = df[df['Contract ID'].notnull()]
df['file_date'] = extracted_char
df['file_date'] = pd.to_datetime(df['file_date'])

#df.to_csv(r"C:\Users\p723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\Patient List\RxAnte_patient_list_20230607_ALL_MAPD.csv", index = False)

#df['Contract ID'].value_counts()

df.head()

filtered_columns = [col for col in df.columns if not col.startswith('Poly')]
df_2 = df[filtered_columns]

print("file is read now to the system")
#take a chunk of random 10000 rows
#sample_df = df_2.sample(n=10000, random_state= 42)


#sample_df.dtypes

#converts to date
#sample_df["Diabetes Medications_Index Date"] = pd.to_datetime(df["Diabetes Medications_Index Date"], errors= 'coerce')

#calc end of year
#end_of_year = pd.to_datetime(sample_df["Diabetes Medications_Index Date"].dt.year,format = '%Y') + pd.offsets.YearEnd(0)

#calc allowed gaps days
#sample_df['Diab Allowed gap days'] =round(end_of_year.sub(sample_df["Diabetes Medications_Index Date"]).dt.days*0.2,0)

#Gaps days imcurred
#sample_df['Diab Gap Days Incurred'] = end_of_year.sub(sample_df["Diabetes Medications_Index Date"]).dt.days * (1- pd.to_numeric(sample_df['Diabetes Medications_PDC (YTD)'],errors='coerce'))

#Gap days remaining
#sample_df['Diab Gap Days remaining'] = sample_df['Diab Allowed gap days'] - sample_df['Diab Gap Days Incurred']

#Gap days percent remianing
#sample_df['Diab percent Gap days remaining'] = sample_df['Diab Gap Days remaining']/ sample_df['Diab Allowed gap days']

#Just viewing the dataframe
#sample_df[['Diab Gap Days Incurred','Diab Allowed gap days','Diab Gap Days remaining','Diab percent Gap days remaining','Diabetes Medications_Index Date', 'Diabetes Medications_PDC (YTD)']]

print("Adding gap days")

def add_gapdays_columns(dataframe, column_prefix, column_prefix_alias):
    ''' Add a group of columns related to gap days'''
    dataframe[f"{column_prefix}Index Date"] = pd.to_datetime(dataframe[f"{column_prefix}Index Date"], errors= 'coerce')
    end_of_year = pd.to_datetime(dataframe[f"{column_prefix}Index Date"].dt.year,format = '%Y') + pd.offsets.YearEnd(0)
    dataframe[f'{column_prefix_alias} Allowed gap days'] =round(end_of_year.sub(dataframe[f"{column_prefix}Index Date"]).dt.days*0.2,0) 
    dataframe[f'{column_prefix_alias} Gap Days Incurred'] = dataframe['file_date'].sub(dataframe[f"{column_prefix}Index Date"]).dt.days * (1- pd.to_numeric(dataframe[f'{column_prefix}PDC (YTD)'],errors='coerce'))
    dataframe[f'{column_prefix_alias} Gap Days remaining'] = dataframe[f'{column_prefix_alias} Allowed gap days'] - dataframe[f'{column_prefix_alias} Gap Days Incurred']
    dataframe[f'{column_prefix_alias} percent Gap days remaining'] = dataframe[f'{column_prefix_alias} Gap Days remaining']/ dataframe[f'{column_prefix_alias} Allowed gap days']
    return dataframe

final_df_gap_days = add_gapdays_columns(df_2, 'Diabetes Medications_', 'Diab')
final_df_gap_days = add_gapdays_columns(final_df_gap_days, 'RASA_', 'RAS')
final_df_gap_days = add_gapdays_columns(final_df_gap_days, 'Statins_', 'Statin')

final_df_gap_days.columns

print("Gap days added")



print("Reading member coverage file")

member_coverage_df = pd.read_csv(r"C:\Users\e723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\Member_coverage_Jan_June.csv")
member_coverage_df.columns
len(member_coverage_df)
member_coverage_df = member_coverage_df.drop_duplicates(subset=['health_insurance_benefit_medicare_number'], keep='first')
len(member_coverage_df)

# mem_attributes = pd.merge(mem_attributes,member_coverage_df,left_on='MBI', right_on = 'health_insurance_benefit_medicare_number', how='left')

# mem_attributes[mem_attributes['MBI'] == '4VU3QX7KD16']
# mem_attributes[mem_attributes['MBI'] == '1VT4QV9KE96']
# mem_attributes[mem_attributes['MBI'] == '8EG0VA0WC45']
# mem_attributes[mem_attributes['MBI'] == '1CV9WC9NY54']

# #89338799741

# mem_attributes['contract_number'] = mem_attributes['contract_number'].str.strip()
final_df_gap_days['Member Id'] = final_df_gap_days['Member Id'].str.strip()



############################################# Add MBI to the dataset #############################
print("Adding MBI")


final_df_gap_days['dob_num'] = final_df_gap_days['Member Date of Birth'].astype(str)

final_df_gap_days['dob_num'] = final_df_gap_days['dob_num'].str.replace('-','', regex=True)
final_df_gap_days['Member Id_9'] = final_df_gap_days['Member Id'].str[:9]
final_df_gap_days['Member Id_9'] = final_df_gap_days['Member Id_9'].str.strip()
final_df_gap_days['Member Id_9'] = final_df_gap_days['Member Id_9'].astype(str)

final_df_gap_days['dob_num'] = final_df_gap_days['dob_num'].astype(str)
final_df_gap_days['dob_num'] = final_df_gap_days['dob_num'].str.strip()


member_coverage_df.columns 
member_coverage_df_2 = member_coverage_df[['member_birth_date','contract_number', 'health_insurance_benefit_medicare_number']]
member_coverage_df_2['contract_number'] = member_coverage_df_2['contract_number'].astype(str)
member_coverage_df_2['contract_number'] = member_coverage_df_2['contract_number'].str.strip()
member_coverage_df_2['member_birth_date'] = member_coverage_df_2['member_birth_date'].astype(str)
member_coverage_df_2['member_birth_date'] = member_coverage_df_2['member_birth_date'].str.strip()


final_df_gap_days_2 = pd.merge(final_df_gap_days,member_coverage_df_2,how='left', left_on=['Member Id_9', 'dob_num'], right_on=['contract_number','member_birth_date'])
#final_df_with_fr_prov_w_landmark_mbi['health_insurance_benefit_medicare_number'].isna().value_counts()

final_df_gap_days_2.columns
final_df_gap_days_2.drop(columns=['member_birth_date', 'contract_number'], inplace= True)

#final_df_gap_days_2.health_insurance_benefit_medicare_number.isna().value_counts()



############################################# Add PO to the dataset #############################
print("Adding PO information from Member attributes file")


def find_latest_bcbsm_member_attribute_file(root_folder):
    # Create a list to store the matching file paths
    matching_files = []
    
    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".XLSX") and "BCBSM_MEMBER" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    matching_file = matching_files[-1]
    return matching_file

def find_latest_bcn_member_attribute_file(root_folder):
    # Create a list to store the matching file paths
    matching_files = []
    
    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".XLSX") and "BCN_MEMBER" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    matching_file = matching_files[-1]
    return matching_file



member_attr_folder = r'C:\Users\e723999\Blue Cross Blue Shield of Michigan\Jiang, Yifan - 18777 MI Stars Member Attributes Monthly Analysis\2023'
bcbsm_member_attr_file =  find_latest_bcbsm_member_attribute_file(member_attr_folder)
print(bcbsm_member_attr_file)

bcn_member_attr_file =  find_latest_bcn_member_attribute_file(member_attr_folder)
print(bcn_member_attr_file)


bcbsm_mem_attributes = pd.read_excel(bcbsm_member_attr_file)
bcn_mem_attributes = pd.read_excel(bcn_member_attr_file)


bcbsm_mem_attributes.columns = bcbsm_mem_attributes.columns.str.strip().str.upper()
bcn_mem_attributes.columns = bcn_mem_attributes.columns.str.strip().str.upper()

mem_attributes = pd.concat([bcbsm_mem_attributes, bcn_mem_attributes],ignore_index=True, sort=False)

mem_attributes.columns
final_df_gap_days_2.columns

final_df_gap_days_3 = pd.merge(final_df_gap_days_2,mem_attributes,how='left', left_on=['health_insurance_benefit_medicare_number'], right_on=['MBI'])

final_df_gap_days_3.drop(columns=['EPI', 'MEMBERID', 'CONTRACT_NUMBER', 'MBI', 'GENDER', 'DOB','PAYMT_ADJUSTMT_MONTH', 'ORG_ESRD',
       'HOSPICE', 'NEW_MEMBER', 'MEMBER_SINCE_DATE_2021', 'PRODUCT', 'GROUPNBR' ], inplace= True)



############################################## Integrating Full risk member ##################################
print("Adding full risk provider information from member roaster")


def find_latest_roster_files(root_folder):
    # Create a list to store the matching file paths
    matching_files = []
    
    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".csv") and "Roster" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    
    return matching_files



roster_root = r'Z:\Outgoing_Reports\Member_Roster'
member_roster_files =  find_latest_roster_files(roster_root)
print(member_roster_files)

latest_roster_file = member_roster_files[-1]
print(f' Loading....{latest_roster_file}')

member_roster_df = pd.read_csv(latest_roster_file, sep='|')
member_roster_df.head()
member_roster_df = member_roster_df[member_roster_df['outcome_type'] == 'Primary']
member_roster_df.columns
member_roster_mini_cols = ['medicare_beneficiary_identifier',  'program_code', 'program_name', 'primary_care_provider_national_provider_identifier','primary_care_provider_name', 'primary_care_provider_group_name']
member_roster_mini_df = member_roster_df[member_roster_df.columns.intersection(member_roster_mini_cols)] 


final_df_gap_days_4 = pd.merge(final_df_gap_days_3,member_roster_mini_df,how='left', left_on=['health_insurance_benefit_medicare_number'], right_on=['medicare_beneficiary_identifier'])


final_df_gap_days_4 = final_df_gap_days_4.drop_duplicates(subset=['Member Id'], keep='first')
len(final_df_gap_days_4)
final_df_gap_days_4.drop(columns= ['medicare_beneficiary_identifier'], inplace=True)

##### Integrating Oakstreet health #######

oakstreet_membership_df = pd.read_excel(r"Z:\Oakstreet\Oak_Street_October_2023_Membership.xlsx", sheet_name='Oak Street October Membership')
oakstreet_membership_df.head()
oakstreet_membership_df = oakstreet_membership_df[['mbi_number', 'pcp_npi']]
oakstreet_membership_df.rename(columns={"pcp_npi": "oakstreet_pcp_npi"}, inplace=True)

final_df_gap_days_5 = pd.merge(final_df_gap_days_4,oakstreet_membership_df,how='left', left_on=['health_insurance_benefit_medicare_number'], right_on=['mbi_number'])
final_df_gap_days_5.drop(columns=['mbi_number'], inplace=True)
final_df_gap_days_5 = final_df_gap_days_5.drop_duplicates(subset=['Member Id'], keep='first')

#final_df_gap_days_5.head()
#final_df_gap_days_5['final_program_name'] = np.where(final_df_gap_days_5['program_name'].isnull(), final_df_gap_days_5['PO_NAME'], final_df_gap_days_5['program_name'] )
final_df_gap_days_5['final_program_name'] = final_df_gap_days_5['program_name']


print("full risk provider information from member roaster added")


########################################### Last year Non-adherent ##########################

print("Adding last year non adherent")


last_year_non_adh_df_PPO = pd.read_excel(r"C:\Users\e723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\From Erica\Part D Ad Hoc Request 06_13_23.xlsx", sheet_name='MAPPO_Part D') 

last_year_non_adh_df_HMO = pd.read_excel(r"C:\Users\e723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\From Erica\Part D Ad Hoc Request 06_13_23.xlsx", sheet_name='BCNA_Part D') 

last_year_non_adh_df_HMO.columns
last_year_non_adh_df_PPO.columns

last_year_non_adh_df_HMO = last_year_non_adh_df_HMO.drop(columns=['MEMBERKEY'])


last_year_non_adh_df_HMO = last_year_non_adh_df_HMO.rename(columns={'ALT_MEMBER_ID_CONTRACT': "MEMBERKEY"})

last_year_non_adh_df_HMO['Org'] = 'HMO'
last_year_non_adh_df_PPO['Org'] = 'PPO'

last_year_non_adh_df_PPO.columns = last_year_non_adh_df_HMO.columns


last_year_non_adh_df = pd.concat([last_year_non_adh_df_HMO, last_year_non_adh_df_PPO], ignore_index= True, sort = False)
last_year_non_adh_df = last_year_non_adh_df[last_year_non_adh_df['NUMERCNT']==0]

last_year_non_adh_df['MEMBERKEY'] = last_year_non_adh_df['MEMBERKEY'].astype(str)
last_year_non_adh_df['MEMBERKEY'] = last_year_non_adh_df['MEMBERKEY'].str.strip()


last_year_non_adh_df_diab = last_year_non_adh_df[last_year_non_adh_df['MEASUREKEY']=='MA-D'][['MEMBERKEY','DENOMCNT','Org']]
last_year_non_adh_df_stat = last_year_non_adh_df[last_year_non_adh_df['MEASUREKEY']=='MA-C'][['MEMBERKEY','DENOMCNT','Org']]
last_year_non_adh_df_hyper = last_year_non_adh_df[last_year_non_adh_df['MEASUREKEY']=='MA-H'][['MEMBERKEY','DENOMCNT','Org']]

final_df_gap_days_5['Member Id'] = final_df_gap_days_5['Member Id'].str.strip()

# last_year_non_adh_df_diab.to_csv(r'C:\Users\e723999\Downloads\last_year_non_adh_df_diab.csv')
# last_year_non_adh_df_stat.to_csv(r'C:\Users\e723999\Downloads\last_year_non_adh_df_stat.csv')
# last_year_non_adh_df_hyper.to_csv(r'C:\Users\e723999\Downloads\last_year_non_adh_df_rasa.csv')

final_df_w_meaningful_risk_nadh = pd.merge(final_df_gap_days_5,last_year_non_adh_df_hyper,how='left', left_on='Member Id', right_on='MEMBERKEY')
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.rename(columns={'DENOMCNT': "Prior_year_NON_Adherent_Hypertention"})
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.drop(columns=['MEMBERKEY','Org'])

final_df_w_meaningful_risk_nadh = pd.merge(final_df_w_meaningful_risk_nadh,last_year_non_adh_df_stat,how='left', left_on='Member Id', right_on='MEMBERKEY')
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.rename(columns={'DENOMCNT': "Prior_year_NON_Adherent_Statin"})
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.drop(columns=['MEMBERKEY','Org'])


final_df_w_meaningful_risk_nadh = pd.merge(final_df_w_meaningful_risk_nadh,last_year_non_adh_df_diab,how='left', left_on='Member Id', right_on='MEMBERKEY')
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.rename(columns={'DENOMCNT': "Prior_year_NON_Adherent_daibetes"})
final_df_w_meaningful_risk_nadh = final_df_w_meaningful_risk_nadh.drop(columns=['MEMBERKEY', 'Org'])

#final_df_w_meaningful_risk_nadh[final_df_w_meaningful_risk_nadh['Member Id'] == '842003088']

#last_year_non_adh_df_hyper[last_year_non_adh_df_hyper['MEMBERKEY']== '842003088']


###################################################### RBCE_MASTER_ROSTER #########################################

print("adding sub-rbce information")
#the below file needs to be updated on a weekly basis

rbce_master = r"C:\Users\e723999\Blue Cross Blue Shield of Michigan\Blueprint for Affordability - Affiliation-Snapshot Practitioner Files\2023\December 31, 2023\RBCE_MASTER_ROSTER_2023_202312_20240101.csv"

fr_providers_df = pd.read_csv(rbce_master) 
fr_providers_df.dtypes

fr_providers_df =  fr_providers_df[(fr_providers_df['Practitioner Affiliation End Date'].isin( ['01/01/4000', '1/1/4000'])) & (fr_providers_df['Network'] == 'FR_MAPPO')][['Practitioner NPI1','Network','Sub-RBCE ID','PU ID', 'PU Name']]


final_df_with_fr_prov = pd.merge(final_df_w_meaningful_risk_nadh,fr_providers_df,how='left', left_on='PCP_NPI', right_on='Practitioner NPI1')




########################################## Adding Landmark Indicator #################################################

print("Adding Landmark Indicator for Landmark engaged members")

landmark_df_unfiltered = member_roster_df[['member_identifier','outcome_description','contract_arrangement_name']]

landmark_df = landmark_df_unfiltered[(landmark_df_unfiltered['outcome_description'].isin(['Case Rate', 'Full Risk'])) & (landmark_df_unfiltered['contract_arrangement_name'] == 'Landmark') ]

#landmark_df.head()
landmark_df['Landmark_indicator'] = 'Y'
landmark_df =  landmark_df[['member_identifier', 'Landmark_indicator' ]]
landmark_df['member_id_str'] = landmark_df['member_identifier']

landmark_df['member_id_str'] = landmark_df['member_identifier'].astype(str)
landmark_df['member_id_str'] = landmark_df['member_id_str'].str.strip()


final_df_with_fr_prov_w_landmark = pd.merge(final_df_with_fr_prov,landmark_df,how='left', left_on='Member Id', right_on='member_id_str')


#final_df_with_fr_prov_w_landmark['Landmark_indicator'].value_counts()

final_df_with_fr_prov_w_landmark = final_df_with_fr_prov_w_landmark.drop(['member_identifier', 'member_id_str'], axis=1)



final_df_with_fr_prov_w_landmark.head(5)


#final_df_with_fr_prov_w_landmark_mbi[final_df_with_fr_prov_w_landmark_mbi['health_insurance_benefit_medicare_number'] == '1CV9WC9NY54']

#add a gap days Less than flag
final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'] = final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'].fillna(9999)
final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'] = final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'].fillna(9999)
final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'] = final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'].fillna(9999)

final_df_with_fr_prov_w_landmark['gap_days_less_than_n'] = np.where((final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'] < 15) & (final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'] > 2) | (final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'] < 15) & (final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'] > 2) | (final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'] < 15) & (final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'] > 2), 1, 0 )

final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'] = np.where(final_df_with_fr_prov_w_landmark['Diab Gap Days remaining']== 9999, np.nan,final_df_with_fr_prov_w_landmark['Diab Gap Days remaining'])
final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'] = np.where(final_df_with_fr_prov_w_landmark['RAS Gap Days remaining']== 9999, np.nan,final_df_with_fr_prov_w_landmark['RAS Gap Days remaining'])
final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'] = np.where(final_df_with_fr_prov_w_landmark['Statin Gap Days remaining']== 9999, np.nan, final_df_with_fr_prov_w_landmark['Statin Gap Days remaining'])

#final_df_with_fr_prov_w_landmark[['Diab Gap Days remaining','RAS Gap Days remaining','Statin Gap Days remaining','gap_days_less_than_n' ]].head(350)

print('Data run complete. Final table ready to export')


############## Incorporating Arine Ind #########################

# non_trust_arine = pd.read_csv(r"W:\STARS_2023\Stars Team\Akshay\PDC - Trust\Arine List\Non trust High Risk members_9_25_2023.csv")
# trust_arine = pd.read_csv(r"W:\STARS_2023\Stars Team\Akshay\PDC - Trust\Arine List\Trust High Risk members_9_25_2023.csv")

# non_trust_arine['Trust ind'] = "Non-Trust" 
# trust_arine['Trust ind'] = "Trust" 
# arine_list_concat = pd.concat([non_trust_arine, trust_arine] ).reset_index()

# arine_list_concat['Trust_measure'] = arine_list_concat['Measure'] + "_" + arine_list_concat['Trust ind']

# arine_list_concat.drop(columns=['index', 'RACE_Cat', 'Neighborhood stress Tier',
#        'Disability', 'Measure', 'Trust ind'], inplace = True)


# arine_list_concat['Count_mem'] = 1

# arine_list_concat_pv = arine_list_concat.pivot_table(index= ['Member Id'], columns='Trust_measure', values='Count_mem', aggfunc='first',).reset_index()

# arine_list_concat_pv['Member Id'] = arine_list_concat_pv['Member Id'].astype(str).str.strip()
# final_df_with_fr_prov_w_landmark['Member Id'] = final_df_with_fr_prov_w_landmark['Member Id'].astype(str).str.strip()

# final_df_with_fr_prov_w_landmark_w_arine  = pd.merge(final_df_with_fr_prov_w_landmark,arine_list_concat_pv, how = 'left', on= "Member Id" )

# final_df_with_fr_prov_w_landmark_w_arine.to_excel(rf"W:\STARS_2023\Stars Team\Akshay\PDC - Trust\Patient list w Arine\RxAnte_patient_list_{extracted_char}_modified__AP.xlsx", index=False)




#################################################### final export of the file ##########################################

print("Exporting the Patient list file")
final_df_with_fr_prov_w_landmark.to_excel(rf"W:\STARS_2023\Stars Team\Akshay\Patient Lists\RxAnte_patient_list_{extracted_char}_modified__AP.xlsx", index=False)
final_df_with_fr_prov_w_landmark.to_csv(rf"W:\STARS_2023\Stars Team\Akshay\OW transitions\BCBSM Rx Stars_Python analysis\RxAnte\RxAnte_patient_list_{extracted_char}_modified__AP.csv", index=False)

#final_df_with_fr_prov_w_landmark_mbi.to_excel(rf"C:\Users\e723999\Downloads\RxAnte_patient_list_{extracted_char}_modified_AP.xlsx", index=False)

print("Patient list export is now complete")

##################################### add diabetes data #####################################

def find_daib_med_files(root_folder):
    # Create a list to store the matching file paths
    matching_files = []
    
    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".xlsx") and "MR_ADH" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    
    return matching_files



daib_med_root = r'W:\STARS_2023\Stars Team\Akshay\Meaningful Risk\Diabetes Roster'
member_daib_med_files =  find_daib_med_files(daib_med_root)
print(member_daib_med_files)

latest_daib_med_file = member_daib_med_files[-1]
print(f' Loading....{latest_daib_med_file}')

member_diab_med_df = pd.read_excel(latest_daib_med_file)

member_diab_med_df.columns
member_diab_med_df  = member_diab_med_df[['Member ID','On Brand Name DIA Meds Only']]

final_df_with_fr_prov_w_landmark =  pd.merge(final_df_with_fr_prov_w_landmark, member_diab_med_df, how='left', left_on='Member Id', right_on='Member ID')


final_df_with_fr_prov_w_landmark[final_df_with_fr_prov_w_landmark['Member Id']=='842134908']



###################################################### Adding member address and Phone number ##############################################################

######## PPO file ########


def find_member_address_files(root_folder, file_ext, contains):    
    latest_file = None
    latest_create_time = 0
    for filename in os.listdir(root_folder):
        if filename.endswith(file_ext) and contains in filename and "~$" not in filename:
            file_path = os.path.join(root_folder, filename)
            create_time = os.path.getctime(file_path)
            if create_time > latest_create_time:
                latest_create_time = create_time
                latest_file = file_path
    return latest_file


member_address_root = r'W:\Rx_Ante\Outbound'
latest_mem_add_ppo_files =  find_member_address_files(member_address_root, ".txt", "enrollment")
print(f' Loading....{latest_mem_add_ppo_files}')

######## HMO File ############

latest_mem_add_hmo_files =  find_member_address_files(member_address_root, ".txt", "Enrollment_BCN")
print(f' Loading....{latest_mem_add_hmo_files}')

mem_add_ppo_df = pd.read_csv(latest_mem_add_ppo_files, sep='|')
mem_add_hmo_df = pd.read_csv(latest_mem_add_hmo_files, sep='|')

mem_add_df = pd.concat([mem_add_ppo_df, mem_add_hmo_df], sort = False)
mem_add_df = mem_add_df[['Member_ID','End_Date','Hic_Number','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number','Member_Alt3_Phone_Number', 'Member_Alt4_Phone_Number', 'Member_Language_Preference']]

mem_add_df.head(10)
mem_add_df.dtypes

mem_add_df['End_Date'] = pd.to_datetime(mem_add_df['End_Date'])

latest_indices = mem_add_df.groupby('Member_ID')['End_Date'].idxmax()
mem_add_df_latest = mem_add_df.loc[latest_indices].reset_index()

mem_add_df_latest['Member_ID'] = mem_add_df_latest['Member_ID'].astype(str).str.strip()
mem_add_df_latest['Member_ID'] = mem_add_df_latest['Member_ID'].str.rstrip('.0')

mem_add_df_latest.drop_duplicates(subset='Member_ID', keep='last', inplace = True)


final_df_with_fr_prov_w_landmark_w_add =  pd.merge(final_df_with_fr_prov_w_landmark, mem_add_df_latest, how='left', left_on='Member Id', right_on='Member_ID')



################################################## PO Lists #######################################

print("exporting the PO lists")

final_df_with_fr_prov_w_landmark_w_add['final_program_name'].value_counts()
final_df_with_fr_prov_w_landmark_w_add['Contract ID'] = final_df_with_fr_prov_w_landmark_w_add['Contract ID'].astype(str)
final_df_with_fr_prov_w_landmark_w_add['Contract ID'] = final_df_with_fr_prov_w_landmark_w_add['Contract ID'].str.strip()

final_df_with_fr_prov_w_landmark_gap_ltd = final_df_with_fr_prov_w_landmark_w_add[(final_df_with_fr_prov_w_landmark_w_add['gap_days_less_than_n'] == 1)
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Futile Date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Futile Date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Futile Date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name', 'primary_care_provider_group_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes', 'oakstreet_pcp_npi',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]
final_df_with_fr_prov_w_landmark_gap_ltd.rename(columns={'Diabetes Medications_Futile Date': "Diabetes Medications_Non-recoverable date",
                                         'RASA_Futile Date': "RASA_Non-recoverable date",
                                         'Statins_Futile Date': "Statins_Non-recoverable date"
                                         }, inplace=True)
final_df_with_fr_prov_w_landmark_gap_ltd["Diabetes Medications_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["Diabetes Medications_Non-recoverable date"].astype(str)
final_df_with_fr_prov_w_landmark_gap_ltd["Diabetes Medications_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["Diabetes Medications_Non-recoverable date"].str.strip()

final_df_with_fr_prov_w_landmark_gap_ltd["RASA_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["RASA_Non-recoverable date"].astype(str)
final_df_with_fr_prov_w_landmark_gap_ltd["RASA_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["RASA_Non-recoverable date"].str.strip()

final_df_with_fr_prov_w_landmark_gap_ltd["Statins_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["Statins_Non-recoverable date"].astype(str)
final_df_with_fr_prov_w_landmark_gap_ltd["Statins_Non-recoverable date"] = final_df_with_fr_prov_w_landmark_gap_ltd["Statins_Non-recoverable date"].str.strip()




pt_list_dt = datetime.strptime(extracted_char, '%Y%m%d')
three_weeks = timedelta(days = 3*7) 
pt_list_dt_target = pt_list_dt + three_weeks



final_df_with_fr_prov_w_landmark_gap_ltd['Non-recoverable_3_week Ind'] = np.where((pd.to_datetime(final_df_with_fr_prov_w_landmark_gap_ltd['Diabetes Medications_Non-recoverable date']) < pt_list_dt_target) | (pd.to_datetime(final_df_with_fr_prov_w_landmark_gap_ltd['RASA_Non-recoverable date']) < pt_list_dt_target) | (pd.to_datetime(final_df_with_fr_prov_w_landmark_gap_ltd['Statins_Non-recoverable date']) < pt_list_dt_target), 1, 0 )

final_df_with_fr_prov_w_landmark_gap_ltd = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Non-recoverable_3_week Ind'] == 1)]


final_df_with_fr_prov_w_landmark_gap_ltd['Diab Gap Days remaining'] = round(final_df_with_fr_prov_w_landmark_gap_ltd['Diab Gap Days remaining'],0) 
final_df_with_fr_prov_w_landmark_gap_ltd['RAS Gap Days remaining'] = round(final_df_with_fr_prov_w_landmark_gap_ltd['RAS Gap Days remaining'],0) 
final_df_with_fr_prov_w_landmark_gap_ltd['Statin Gap Days remaining'] = round(final_df_with_fr_prov_w_landmark_gap_ltd['Statin Gap Days remaining'],0) 




final_df_with_fr_prov_w_landmark_gap_diab = final_df_with_fr_prov_w_landmark_w_add[(final_df_with_fr_prov_w_landmark_w_add['On Brand Name DIA Meds Only']== '1')
                                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Futile Date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME', 'oakstreet_pcp_npi', 'primary_care_provider_group_name',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]

final_df_with_fr_prov_w_landmark_gap_diab.rename(columns={'Diabetes Medications_Futile Date': "Diabetes Medications_Non-recoverable date"}, inplace=True)


print("exporting the Landmark files")
final_df_with_fr_prov_w_landmark_gap_ltd
#Name: Landmark	
#delete Sub-RBCE for PPO 
# HMO remove  :  IH0000000011, IH0000000017 ,IH0000000118,IH0000000020,IH0000000123,IH0000000131,IH0000000138,IH0000000140,IH0000000148
#gap days <20 in any 1 drug  - final_df_with_fr_prov_w_landmark_w_add['gap_days_less_than_n'] = 1
# Add landmark engaged only

final_df_landmark_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['final_program_name'] == 'Landmark') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572') &
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'].isna()) &
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Landmark_indicator'] == 'Y') &
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['primary_care_provider_group_name'] !='Oak Street Health') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]

#final_df_landmark_ppo_v0[final_df_landmark_ppo_v0['Member Id'] == '990326665']

final_df_landmark_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['final_program_name'] == 'Landmark') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572') &
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'].isna()) &
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Landmark_indicator'] == 'Y') &
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['primary_care_provider_group_name'] !='Oak Street Health') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]




PO_IDs = ['IH0000000011', 'IH0000000017' ,'IH0000000118','IH0000000020','IH0000000123','IH0000000131','IH0000000138','IH0000000140','IH0000000148']
final_df_landmark_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['final_program_name'] == 'Landmark') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883') &
                                                             ~(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(PO_IDs)) &
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Landmark_indicator'] == 'Y') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['primary_care_provider_group_name'] !='Oak Street Health') 
                                                                                                                         ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_landmark_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['final_program_name'] == 'Landmark') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883') &
                                                             ~(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(PO_IDs)) &
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Landmark_indicator'] == 'Y') &
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['primary_care_provider_group_name'] !='Oak Street Health') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]





excel_file_landmark_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Landmark\Landmark_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_landmark_ppo_path) as writer:
    final_df_landmark_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_landmark_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_landmark_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Landmark\Landmark_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_landmark_hmo_path) as writer:
    final_df_landmark_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_landmark_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)




print("exporting the OPNS")

# Name: Oakland Physician Network Services	
# Sub-RBCE : SRBCE0006015 
# PO_ID : IH0000000011

final_df_OPNS_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006015') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OPNS_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006015') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OPNS_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'] == 'IH0000000011') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OPNS_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'] == 'IH0000000011') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



excel_file_OPNS_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Oakland Physician Network Services\rbce015~OPNS_drug_adherence_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_OPNS_ppo_path) as writer:
    final_df_OPNS_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_OPNS_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_OPNS_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Oakland Physician Network Services\j00as~OPNS_drug_adherence_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_OPNS_hmo_path) as writer:
    final_df_OPNS_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_OPNS_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


print("exporting the United Physicians")

# Name: United Physicians	
# Sub-RBCE : SRBCE0006001 
# PO_ID : IH0000000017 IH0000000118


final_df_UP_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006001') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_UP_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab [(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006001') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572') 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_UP_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(['IH0000000017', 'IH0000000118'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_UP_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(['IH0000000017', 'IH0000000118'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


excel_file_UP_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\United Physicians\rbce001~UP_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_UP_ppo_path) as writer:
    final_df_UP_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_UP_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_UP_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\United Physicians\j000y~UP_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_UP_hmo_path) as writer:
    final_df_UP_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_UP_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


print("exporting the Medical Network One")

# Name: Medical Network One	
# Sub-RBCE : SRBCE0006019 
# PO_ID : IH0000000020


final_df_MNO_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006019') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_MNO_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006019') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_MNO_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(['IH0000000020'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_MNO_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(['IH0000000020'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



excel_file_MNO_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Medical Network One\rbce019~MNO_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_MNO_ppo_path) as writer:
    final_df_MNO_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_MNO_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_MNO_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Medical Network One\j000q~MNO_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_MNO_hmo_path) as writer:
    final_df_MNO_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_MNO_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


print("exporting the Great Lakes OSC")

# Name: Great Lakes OSC
# Sub-RBCE : SRBCE0006003 
# PO_ID : IH0000000131


final_df_GLO_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006003') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_GLO_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006003') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_GLO_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(['IH0000000131'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]

final_df_GLO_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(['IH0000000131'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



excel_file_GLO_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Great Lakes OSC\rbce003~GLOSC_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_GLO_ppo_path) as writer:
    final_df_GLO_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_GLO_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_GLO_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Great Lakes OSC\j00qa~GLOSC_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_GLO_hmo_path) as writer:
    final_df_GLO_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_GLO_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)

print("exporting the Oak Street Health")

# Name: Oak Street Health
# PO_ID : IH0000000138

final_df_OSH_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['primary_care_provider_group_name'] =='Oak Street Health')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['oakstreet_pcp_npi'].notna()) 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OSH_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['primary_care_provider_group_name'] =='Oak Street Health')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['oakstreet_pcp_npi'].notna()) 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


excel_file_OSH_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Oak Street Health\OSH_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_OSH_hmo_path) as writer:
    final_df_OSH_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_OSH_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


#SRBCE0004006
final_df_OSH_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['primary_care_provider_group_name'] =='Oak Street Health')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['oakstreet_pcp_npi'].notna()) 
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OSH_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['primary_care_provider_group_name'] =='Oak Street Health')& 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['oakstreet_pcp_npi'].notna())
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



excel_file_OSH_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Oak Street Health\OSH_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_OSH_ppo_path) as writer:
    final_df_OSH_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_OSH_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



print("exporting the Dedicated")

# Name: Dedicated Physicians Group of MI
# Sub-RBCE : SRBCE0004023 
# PO_ID : IH0000000140

#NO members were found in PPO
final_df_DPG_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0004023') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]

final_df_DPG_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0004023') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_DPG_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(['IH0000000140'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_DPG_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(['IH0000000140'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]




excel_file_DPG_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Dedicated Physicians Group of MI\DPG_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_DPG_ppo_path) as writer:
    final_df_DPG_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_DPG_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_DPG_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Dedicated Physicians Group of MI\DPG_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_DPG_hmo_path) as writer:
    final_df_DPG_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_DPG_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


print("exporting the HVPA")

# Name: Huron Valley Physicians Association
# Sub-RBCE : SRBCE0006029 
# PO_ID : IH0000000148


final_df_HVP_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006029') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_HVP_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006029') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



final_df_HVP_hmo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['PO_ID'].isin(['IH0000000148'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_HVP_hmo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['PO_ID'].isin(['IH0000000148'])) & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H5883')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]



excel_file_HVP_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Huron Valley Physicians Association\rbce029~HVPA_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_HVP_ppo_path) as writer:
    final_df_HVP_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_HVP_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



excel_file_HVP_hmo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\Huron Valley Physicians Association\j00bn~HVPA_drug_adhernce_HMO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_HVP_hmo_path) as writer:
    final_df_HVP_hmo_v0.to_excel(writer, sheet_name='HMO_gap_days', index = False)
    final_df_HVP_hmo_diab.to_excel(writer, sheet_name="diab_medication", index=False)


print("exporting the OSP")

# Name: Oakland Southfield Physicians
# Sub-RBCE : SRBCE0006002 


final_df_OSP_ppo_v0 = final_df_with_fr_prov_w_landmark_gap_ltd[(final_df_with_fr_prov_w_landmark_gap_ltd['Sub-RBCE ID'] == 'SRBCE0006002') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_ltd['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	'Recommended Intervention',	'Previous Intervention',	'Therapies In Play',	'Outreach Therapy',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'RASA_Index Date',	'RASA_New to Therapy',	'RASA_Most Recent Rx',	'RASA_NDC',	'RASA_Most Recent Fill Date',	'RASA_Next Fill Due Date',	'RASA_Pharmacy NPI',	'RASA_Pharmacy Name',	'RASA_Pharmacy Phone',	'RASA_Prescriber NPI',	'RASA_Star Status',	'RASA_Days Supply to Adherent',	'RASA_PDC (YTD)',	'RASA_Reason For Outreach','RASA_Adherence Risk',	'RASA_Current Fill Status',	'RASA_Non-recoverable date',	'RASA_Total Days Supply (YTD)',	'RASA_Fill Count (YTD)',	
                                                                'Statins_Index Date',	'Statins_New to Therapy',	'Statins_Most Recent Rx',	'Statins_NDC',	'Statins_Most Recent Fill Date',	'Statins_Next Fill Due Date',	'Statins_Pharmacy NPI',	'Statins_Pharmacy Name',	'Statins_Pharmacy Phone',	'Statins_Prescriber NPI',	'Statins_Star Status',	'Statins_Days Supply to Adherent','Statins_PDC (YTD)',	'Statins_Reason For Outreach',	'Statins_Adherence Risk',	'Statins_Current Fill Status',	'Statins_Non-recoverable date',	'Statins_Total Days Supply (YTD)',	'Statins_Fill Count (YTD)',	'Diab Gap Days remaining',	'RAS Gap Days remaining',	'Statin Gap Days remaining',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name',	'Prior_year_NON_Adherent_Hypertention',	'Prior_year_NON_Adherent_Statin',	
                                                                'Prior_year_NON_Adherent_daibetes',	'Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','gap_days_less_than_n',
                                                                'Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]


final_df_OSP_ppo_diab = final_df_with_fr_prov_w_landmark_gap_diab[(final_df_with_fr_prov_w_landmark_gap_diab['Sub-RBCE ID'] == 'SRBCE0006002') & 
                                                             (final_df_with_fr_prov_w_landmark_gap_diab['Contract ID'] =='H9572')
                                                             ][['Member Id',	'Contract ID',	'Member First Name',	'Member Last Name',	'Member Date of Birth',	
                                                                'Diabetes Medications_Index Date',	'Diabetes Medications_New to Therapy',	'Diabetes Medications_Most Recent Rx',	'Diabetes Medications_NDC',	'Diabetes Medications_Most Recent Fill Date',	'Diabetes Medications_Next Fill Due Date',	'Diabetes Medications_Pharmacy NPI','Diabetes Medications_Pharmacy Name',	'Diabetes Medications_Pharmacy Phone',	'Diabetes Medications_Prescriber NPI',	'Diabetes Medications_Star Status',	'Diabetes Medications_Days Supply to Adherent',	'Diabetes Medications_PDC (YTD)',	'Diabetes Medications_Reason For Outreach',	'Diabetes Medications_Adherence Risk',	'Diabetes Medications_Current Fill Status',	'Diabetes Medications_Non-recoverable date',	'Diabetes Medications_Total Days Supply (YTD)',	'Diabetes Medications_Fill Count (YTD)',	
                                                                'PO_ID',	'PO_NAME',	'final_program_name',	'PCP_NPI',	'primary_care_provider_name','Sub-RBCE ID',	'PU ID',	'PU Name',	'Landmark_indicator',	'health_insurance_benefit_medicare_number','On Brand Name DIA Meds Only','Member_Address_1',	'Member_Address_2',	'Member_City','Member_State',	'Member_Zip_Code',	'Member_Country_Code',	'Member_Eligibility_Phone_Number', 	'Member_Phone_Number',	'Member_Preferred_Phone_Number','Member_Alt1_Phone_Number',	'Member_Alt2_Phone_Number', 'Member_Language_Preference']]




excel_file_OSP_ppo_path = rf"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists\OSP\rbce002~OSP_drug_adhernce_PPO_{extracted_char}_list.xlsx"

with pd.ExcelWriter(excel_file_OSP_ppo_path) as writer:
    final_df_OSP_ppo_v0.to_excel(writer, sheet_name='PPO_gap_days', index = False)
    final_df_OSP_ppo_diab.to_excel(writer, sheet_name="diab_medication", index=False)



print("All PO list files exported")







##################################################################################################
############################## PO PCP Providers ##################################################
##################################################################################################


#reading HMO PCP files

def find_HMO_PCP_files(root_folder):
    matching_files = []
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_list" in the name
            if file.endswith(".xlsx") and "MADV" in file and "~$" not in file :
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Add the file path to the matching_files list
                matching_files.append(file_path)
    return matching_files



pcp_hmo_root = r'C:\Users\e723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\PCP info\HMO'
pcp_hmo_files =  find_HMO_PCP_files(pcp_hmo_root)
print(pcp_hmo_files)

latest_pcp_hmo_file = pcp_hmo_files[-1]
print(f' Loading....{latest_pcp_hmo_file}')

member_pcp_hmo_df = pd.read_excel(latest_pcp_hmo_file, sheet_name= 'Member Detail', skiprows=8)
member_pcp_hmo_df.head()

member_pcp_hmo_df.columns
member_pcp_hmo_df.dtypes

member_pcp_hmo_df['MEMBER #'] = member_pcp_hmo_df['MEMBER #'].astype(str).str.strip()



#reading PPO PCP files


def find_PPO_PCP_files(root_folder, file_ext, contains):
    # Create a variable to store the latest file path
    latest_file_path = None
    latest_creation_time = 0

    # Traverse through all subfolders and files in the root folder
    for root, dirs, files in os.walk(root_folder):
        for file in files:
            # Check if the file is an Excel file (.xlsx) and contains "patient_1"
            if file.endswith(file_ext) and contains in file and "$" not in file:
                # Build the absolute file path
                file_path = os.path.join(root, file)
                
                # Get the creation time of the file
                file_creation_time = os.path.getctime(file_path)

                # Check if it's the latest file
                if file_creation_time > latest_creation_time:
                    latest_creation_time = file_creation_time
                    latest_file_path = file_path

    # Return the latest file path
    return latest_file_path




pcp_ppo_root = r'C:\Users\e723999\OneDrive - Blue Cross Blue Shield of Michigan\PDC files\Diabetes_deepdive\PCP info\PPO'
pcp_ppo_files =  find_PPO_PCP_files(pcp_ppo_root, ".xlsx", "MbrAttr")
print(pcp_ppo_files)


print(f' Loading....{pcp_ppo_files}')

member_pcp_ppo_df = pd.read_excel(pcp_ppo_files)
#member_pcp_ppo_df.head()
#member_pcp_ppo_df.columns


member_pcp_ppo_df['PCP NAME'] =  member_pcp_ppo_df['PractitionerLastName'].str.cat(member_pcp_ppo_df['PractitionerFirstName'],sep=", ")


member_pcp_hmo_df = member_pcp_hmo_df.rename(columns={'MCG': 'PO_ID', "NAME": "OrgTitle", 'PCP NPI': 
                                  'NPI', 'MEMBER #': 'ContractNum'})

member_pcp_ppo_df = member_pcp_ppo_df[['ContractNum','PO_ID','OrgTitle','SubPO_ID','SubGroup','PU_ID', 'PracticeUnit','PCP NAME','NPI']]
member_pcp_hmo_df = member_pcp_hmo_df[['PO_ID', 'OrgTitle', 'PCP NAME', 'NPI','ContractNum']]


#concating the HMO and PPO file
final_pcp_df = pd.concat([member_pcp_ppo_df,member_pcp_hmo_df], sort = True).reset_index()



#member_pcp_hmo_df[(member_pcp_hmo_df['ContractNum']=='99701006201')]   

final_pcp_df['ContractNum'] = final_pcp_df['ContractNum'].astype(str).str.strip()
final_pcp_df['PO_ID'].value_counts()

#final_pcp_df[(final_pcp_df['ContractNum']=='99701006201')]   

#READING REFILLS REMAINING FILE

member_adh_med_df = pd.read_excel(latest_daib_med_file)

member_adh_med_df.head()
member_adh_med_df.drop(columns= ['On Brand Name DIA Meds Only'], inplace = True)


final_df_gap_days['Member Id'] = final_df_gap_days['Member Id'].astype(str).str.strip()

member_adh_med_df.head()

# Merging PCP information with RxAnte Patient list
pcp_with_gap_days = pd.merge(final_df_gap_days,final_pcp_df,left_on='Member Id', right_on = 'ContractNum', how='left')

# pcp_with_gap_days[(pcp_with_gap_days['ContractNum']=='99701006201')] 
# final_df_gap_days[(final_df_gap_days['Member Id']=='99701006201')] 

#Merging the PCP and RxAnte Patient list with Refill remaining. 
final_pcp_with_gap_days_and_ref_rem = pd.merge(pcp_with_gap_days,member_adh_med_df,left_on='Member Id', right_on = 'Member ID', how='left')


final_pcp_with_gap_days_and_ref_rem.columns


final_pcp_with_gap_days_and_ref_rem.drop(columns =['Member ID','ContractNum','Member Id_9','index'], inplace = True)


#Less than 30% gap days
final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'] = final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'].fillna(9999)
final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'] = final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'].fillna(9999)
final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'] = final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'].fillna(9999)

final_pcp_with_gap_days_and_ref_rem['gap_days_less_than_n'] = np.where((final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'] < 0.3) & (final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'] > 0) |
                                                                        (final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'] < 0.3) & (final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'] > 0) | 
                                                                        (final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'] < 0.3) & (final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'] > 0), 1, 0 )

final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'] = np.where(final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining']== 9999, np.nan,final_pcp_with_gap_days_and_ref_rem['Diab percent Gap days remaining'])
final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'] = np.where(final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining']== 9999, np.nan,final_pcp_with_gap_days_and_ref_rem['RAS percent Gap days remaining'])
final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'] = np.where(final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining']== 9999, np.nan, final_pcp_with_gap_days_and_ref_rem['Statin percent Gap days remaining'])


############### adding member address and phone number info ########################

final_pcp_w_gap_and_refill_and_mem = pd.merge(final_pcp_with_gap_days_and_ref_rem, mem_add_df_latest, how = 'left' , left_on = ['Member Id'], right_on=['Member_ID'])



############### Add OW's Need Fill or Need PRescription logic ############################


#Futile date identification logic Tested on multiple date scenario.

# def find_futile_date():
#     today = date.today()
#     first_day_of_next_month = date(today.year, today.month + 1, 1)


#     if first_day_of_next_month.month < 12:

#         # Find the last Wednesday of the current month
#         last_day_of_current_month = first_day_of_next_month - timedelta(days=1)
#         while last_day_of_current_month.weekday() != 2:  # Wednesday is represented as 2
#             last_day_of_current_month -= timedelta(days=1)

#         # Find the first Wednesday of the next month
#         next_month_first_wednesday = first_day_of_next_month
#         while next_month_first_wednesday.weekday() != 2:  # Wednesday is represented as 2
#             next_month_first_wednesday += timedelta(days=1)

#         # Find the first Wednesday of the next to next month

#         next_to_next_month_first_wednesday = date(today.year, today.month + 2, 1)
#         while next_to_next_month_first_wednesday.weekday() != 2:  # Wednesday is represented as 2
#             next_to_next_month_first_wednesday += timedelta(days=1)

#         # Check if today is between the last Wednesday of the current month and the first Wednesday of the next month
#         if today >= last_day_of_current_month and today < next_month_first_wednesday:
#             target_date = next_to_next_month_first_wednesday
#             return target_date
#         elif today < next_month_first_wednesday:
#             target_date = next_month_first_wednesday
#             return target_date
#     else : return date(2023,12,31)
    

futile_by_date = date(2023,12,31)

# Diabetes


final_pcp_w_gap_and_refill_and_mem['DIA Med Refills Remaining'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['DIA Med Refills Remaining'],errors='coerce')
final_pcp_w_gap_and_refill_and_mem['DIA Med Day Supply'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['DIA Med Day Supply'],errors='coerce') 

final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Next Fill Due Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Next Fill Due Date'], errors = 'coerce')

final_pcp_w_gap_and_refill_and_mem['Diabetes needs new script'] = final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Next Fill Due Date'] + pd.to_timedelta(final_pcp_w_gap_and_refill_and_mem['DIA Med Refills Remaining'] * final_pcp_w_gap_and_refill_and_mem['DIA Med Day Supply'],unit='D')

final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Futile Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Futile Date'], errors = 'coerce')

final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Fill Count (YTD)'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Fill Count (YTD)'], errors = 'coerce')
final_pcp_w_gap_and_refill_and_mem['On Diabetes list'] = np.where((final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Futile Date'] <= pd.Timestamp(futile_by_date)) & 
                                                                  (final_pcp_w_gap_and_refill_and_mem['Diabetes Medications_Fill Count (YTD)'] > 1) ,1,0)
final_pcp_w_gap_and_refill_and_mem['Diabetes cohort'] = np.where(final_pcp_w_gap_and_refill_and_mem['On Diabetes list'] == 1, np.where(final_pcp_w_gap_and_refill_and_mem['Diabetes needs new script'] <= pd.Timestamp(futile_by_date),"Needs prescription", np.where(final_pcp_w_gap_and_refill_and_mem['Diabetes needs new script'] >= pd.Timestamp(futile_by_date),"Needs fill","Needs fill or Needs prescription")),np.NaN)


#final_pcp_w_gap_and_refill_and_mem[['Diabetes Medications_Next Fill Due Date','DIA Med Refills Remaining','DIA Med Day Supply','Diabetes needs new script','Diabetes Medications_Futile Date','On Diabetes list','Diabetes cohort' ]]

# RASA

final_pcp_w_gap_and_refill_and_mem['HYP Med Refills Remaining'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['HYP Med Refills Remaining'],errors='coerce')
final_pcp_w_gap_and_refill_and_mem['HYP Med Day Supply'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['HYP Med Day Supply'],errors='coerce') 

final_pcp_w_gap_and_refill_and_mem['RASA_Next Fill Due Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['RASA_Next Fill Due Date'], errors = 'coerce')

final_pcp_w_gap_and_refill_and_mem['RASA needs new script'] = final_pcp_w_gap_and_refill_and_mem['RASA_Next Fill Due Date'] + pd.to_timedelta(final_pcp_w_gap_and_refill_and_mem['HYP Med Refills Remaining'] * final_pcp_w_gap_and_refill_and_mem['HYP Med Day Supply'],unit='D')

final_pcp_w_gap_and_refill_and_mem['RASA_Futile Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['RASA_Futile Date'], errors = 'coerce')


final_pcp_w_gap_and_refill_and_mem['RASA_Fill Count (YTD)'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['RASA_Fill Count (YTD)'], errors = 'coerce')
final_pcp_w_gap_and_refill_and_mem['On RASA list'] = np.where((final_pcp_w_gap_and_refill_and_mem['RASA_Futile Date'] <= pd.Timestamp(futile_by_date)) & 
                                                                  (final_pcp_w_gap_and_refill_and_mem['RASA_Fill Count (YTD)'] > 1) ,1,0)

final_pcp_w_gap_and_refill_and_mem['RASA cohort'] = np.where(final_pcp_w_gap_and_refill_and_mem['On RASA list'] == 1, np.where(final_pcp_w_gap_and_refill_and_mem['RASA needs new script'] <= pd.Timestamp(futile_by_date),"Needs prescription", np.where(final_pcp_w_gap_and_refill_and_mem['RASA needs new script'] >= pd.Timestamp(futile_by_date),"Needs fill","Needs fill or Needs prescription")),np.NaN)


#Statins

final_pcp_w_gap_and_refill_and_mem['CHO Med Refills Remaining'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['CHO Med Refills Remaining'],errors='coerce')
final_pcp_w_gap_and_refill_and_mem['CHO Med Day Supply'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['CHO Med Day Supply'],errors='coerce') 

final_pcp_w_gap_and_refill_and_mem['Statins_Next Fill Due Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['Statins_Next Fill Due Date'], errors = 'coerce')

final_pcp_w_gap_and_refill_and_mem['Statins needs new script'] = final_pcp_w_gap_and_refill_and_mem['Statins_Next Fill Due Date'] + pd.to_timedelta(final_pcp_w_gap_and_refill_and_mem['CHO Med Refills Remaining'] * final_pcp_w_gap_and_refill_and_mem['CHO Med Day Supply'],unit='D')

final_pcp_w_gap_and_refill_and_mem['Statins_Futile Date'] = pd.to_datetime(final_pcp_w_gap_and_refill_and_mem['Statins_Futile Date'], errors = 'coerce')



final_pcp_w_gap_and_refill_and_mem['Statins_Fill Count (YTD)'] = pd.to_numeric(final_pcp_w_gap_and_refill_and_mem['Statins_Fill Count (YTD)'], errors = 'coerce')
final_pcp_w_gap_and_refill_and_mem['On Statins list'] = np.where((final_pcp_w_gap_and_refill_and_mem['Statins_Futile Date'] <= pd.Timestamp(futile_by_date)) & 
                                                                  (final_pcp_w_gap_and_refill_and_mem['Statins_Fill Count (YTD)'] > 1) ,1,0)

final_pcp_w_gap_and_refill_and_mem['Statins cohort'] = np.where(final_pcp_w_gap_and_refill_and_mem['On Statins list'] == 1, np.where(final_pcp_w_gap_and_refill_and_mem['Statins needs new script'] <= pd.Timestamp(futile_by_date),"Needs prescription", np.where(final_pcp_w_gap_and_refill_and_mem['Statins needs new script'] >= pd.Timestamp(futile_by_date),"Needs fill","Needs fill or Needs prescription")),np.NaN)


final_pcp_w_gap_and_refill_and_mem.head()
final_pcp_w_gap_and_refill_and_mem.drop(columns= ['index','Member_ID','End_Date'], inplace = True)


#renaming 
                               
final_pcp_w_gap_and_refill_and_mem.rename(columns={'DIA Med Name': 'Diabetes Med Name',
                                         'DIA Med Refills Remaining': 'Diabetes Med Refills Remaining',
                                         'DIA Med Day Supply': 'Diabetes Med Day Supply',
                                          'HYP Med Name' : 'RASA Med Name',
                                          'HYP Med Refills Remaining' : 'RASA Med Refills Remaining',
                                          'HYP Med Day Supply' : 'RASA Med Day Supply',
                                          'CHO Med Name' : 'Statins Med Name' ,	
                                          'CHO Med Refills Remaining' : 'Statins Med Refills Remaining' ,	
                                          'CHO Med Day Supply' : 'Statins Med Day Supply' 
                                         }, inplace=True)



final_pcp_w_gap_and_refill_and_mem['Member_Language_Preference'] = np.where ((final_pcp_w_gap_and_refill_and_mem['Member_Language_Preference'] == '0')| (final_pcp_w_gap_and_refill_and_mem['Member_Language_Preference'].isna()), "Unknown", final_pcp_w_gap_and_refill_and_mem['Member_Language_Preference'])

final_pcp_w_gap_and_refill_and_mem = final_pcp_w_gap_and_refill_and_mem.replace('nan','')





######################### On a list indicator #######################################

# 1 - Yes, Member on the list for this month

final_pcp_w_gap_and_refill_and_mem['On_a_list_Ind'] = np.where(final_pcp_w_gap_and_refill_and_mem['On Diabetes list']| 
                                                               final_pcp_w_gap_and_refill_and_mem['On RASA list'] | 
                                                               final_pcp_w_gap_and_refill_and_mem['On Statins list'], 1, 0)



# Adding meeting date for POs


def find_latest_matching_files(root_folder, file_ext, contains):    
    latest_file = None
    latest_create_time = 0
    for filename in os.listdir(root_folder):
        if filename.endswith(file_ext) and contains in filename and "~$" not in filename:
            file_path = os.path.join(root_folder, filename)
            create_time = os.path.getctime(file_path)
            if create_time > latest_create_time:
                latest_create_time = create_time
                latest_file = file_path
    return latest_file


root_folder = r"W:\STARS_2023\Stars Team\Akshay\PO_Opt_in_files"
latest_daily_incentive_opt_in_file =  find_latest_matching_files(root_folder, ".xlsx", "daily")
print(f' Loading....{latest_daily_incentive_opt_in_file}')

po_opt_in_list = pd.read_excel(latest_daily_incentive_opt_in_file, sheet_name="OPT Tracking_Incentive by Grp")



po_opt_in_list['PO_MCG ID'] = po_opt_in_list['PO_MCG ID'].astype(str).str.strip()

po_opt_in_list_ltd = po_opt_in_list[['PO_MCG ID', 'Meeting date and time with PO_MCG']]

po_opt_in_list_ltd['Meeting date_PO_MCG'] = np.where(po_opt_in_list_ltd['Meeting date and time with PO_MCG'] =='Scheduling', po_opt_in_list_ltd['Meeting date and time with PO_MCG'].str.split('/', 2).str[:2].str.join('/'),po_opt_in_list_ltd['Meeting date and time with PO_MCG'].str.split('/', 2).str[:2].str.join('/')+"/2023")



final_pcp_w_gap_and_refill_and_mem_opt  = pd.merge(final_pcp_w_gap_and_refill_and_mem, po_opt_in_list_ltd, how = 'left', left_on= 'PO_ID' , right_on= 'PO_MCG ID')

final_pcp_w_gap_and_refill_and_mem_opt.drop(columns=['PO_MCG ID', 'Meeting date and time with PO_MCG'], inplace = True)



#Export master file for tracking

final_pcp_w_gap_and_refill_and_mem_opt.to_csv(rf'W:\STARS_2023\Stars Team\Akshay\PCP Refills remaining\patient_lst_pcp_refill_rem_w_PO_mtg_date_{extracted_char}.csv', index =False)



############## adding the meaningfulrisk member flag ##################


def find_latest_hmo_and_ppo_files(root_dir):
    hmo_latest = {}
    ppo_latest = {}

    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.endswith('.xlsx'):
                file_path = os.path.join(root, file)
                create_time = os.path.getctime(file_path)
                if 'HMO' in file:
                    if root not in hmo_latest or create_time > hmo_latest[root][1]:
                        hmo_latest[root] = (file_path, create_time)
                elif 'PPO' in file:
                    if root not in ppo_latest or create_time > ppo_latest[root][1]:
                        ppo_latest[root] = (file_path, create_time)

    hmo_latest_files = [file for file, _ in hmo_latest.values()]
    ppo_latest_files = [file for file, _ in ppo_latest.values()]

    return hmo_latest_files, ppo_latest_files

def combine_excel_files_to_dataframe(excel_files):
    df_list = []
    for file in excel_files:
        xls = pd.ExcelFile(file)
        if 'PPO_gap_days' in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name='PPO_gap_days')
            df['File'] = os.path.basename(file)  # Add a 'File' column with the file name
            df_list.append(df)
    combined_df = pd.concat(df_list, ignore_index=True)
    return combined_df

# Example usage:
root_directory = r"W:\STARS_2023\Stars Team\Akshay\Outbound_Patient_Lists"
hmo_files, ppo_files = find_latest_hmo_and_ppo_files(root_directory)

combined_MR_df = combine_excel_files_to_dataframe(hmo_files + ppo_files)

combined_MR_df_mem_list  = combined_MR_df[['Member Id']]

combined_MR_df_mem_list['Meaningful_risk_Ind'] = 'Y'

combined_MR_df_mem_list['Member Id'] = combined_MR_df['Member Id'].astype(str).str.strip()

final_pcp_w_gap_and_refill_onlist = final_pcp_w_gap_and_refill_and_mem_opt[(final_pcp_w_gap_and_refill_and_mem_opt['On_a_list_Ind']== 1) &
                                                                        (final_pcp_w_gap_and_refill_and_mem_opt['PO_ID'].notna())]

final_pcp_w_gap_and_refill_onlist['Member Id'] = final_pcp_w_gap_and_refill_onlist['Member Id'].astype(str).str.strip()

final_pcp_w_gap_and_refill_onlist = pd.merge(final_pcp_w_gap_and_refill_onlist, combined_MR_df_mem_list, how = 'left', left_on = 'Member Id', right_on = 'Member Id' )

#final_pcp_w_gap_and_refill_onlist['Meaningful_risk_Ind'].value_counts()


#################### removing data from columns that don't meet the criteria  #########################

                    


final_pcp_w_gap_and_refill_onlist.head()

all_columns_with_daib = []
for i in final_pcp_w_gap_and_refill_onlist.columns:
    if i.startswith("Diabetes"):
        all_columns_with_daib.append(i)

all_columns_with_RASA = []
for i in final_pcp_w_gap_and_refill_onlist.columns:
    if i.startswith("RAS"):
        all_columns_with_RASA.append(i)

all_columns_with_Statin = []
for i in final_pcp_w_gap_and_refill_onlist.columns:
    if i.startswith("Statin"):
        all_columns_with_Statin.append(i)



final_pcp_w_gap_and_refill_onlist.loc[final_pcp_w_gap_and_refill_onlist['On Diabetes list'] == 0 ,all_columns_with_daib] = np.nan
final_pcp_w_gap_and_refill_onlist.loc[final_pcp_w_gap_and_refill_onlist['On RASA list'] == 0 ,all_columns_with_RASA] = np.nan
final_pcp_w_gap_and_refill_onlist.loc[final_pcp_w_gap_and_refill_onlist['On Statins list'] == 0 ,all_columns_with_Statin] = np.nan

final_pcp_w_gap_and_refill_onlist.head()

final_pcp_w_gap_and_refill_onlist = final_pcp_w_gap_and_refill_onlist.replace('nan','')

# final_pcp_w_gap_and_refill_onlist['Diabetes Medications_Star Status'].value_counts()
# final_pcp_w_gap_and_refill_onlist['Diabetes Medications_Reason For Outreach'].value_counts()
# final_pcp_w_gap_and_refill_onlist['Diabetes Medications_Adherence Risk'].value_counts()
# final_pcp_w_gap_and_refill_onlist['Diabetes Medications_Current Fill Status'].value_counts()
# final_pcp_w_gap_and_refill_onlist['Diabetes Medications_Days Supply to Adherent'].value_counts()



### Reduceing and reodering the columns to only rquired ones.

# for i in final_pcp_w_gap_and_refill_onlist.columns:
#     print(i)

#connect edifecs with 

edifecs_file_location = r"W:\STARS_2023\Stars Team\Akshay\PGIP_MCG Edifecs Mailboxes IDs.xlsx"
pgip_ppo_name = pd.read_excel(edifecs_file_location, sheet_name = 'PGIP (MAPPO) ' )

pgip_hmo_name = pd.read_excel(edifecs_file_location, sheet_name = 'MCG (BCNA)' )

pgip_ppo_name['file_name_prefix'] = pgip_ppo_name['EDDI Mailbox'] + "~" + pgip_ppo_name['Mailbox Name']		
pgip_ppo_name_2 = pgip_ppo_name[['PO ID', 'file_name_prefix']]

pgip_hmo_name['file_name_prefix'] = pgip_hmo_name['Mailbox for all other files'] + "~" + pgip_hmo_name['DS']		

pgip_hmo_name.rename(columns = {'ID': 'PO ID'}, inplace = True)

pgip_hmo_name_2 = pgip_hmo_name[['PO ID', 'file_name_prefix']]

pgip_final_name = pd.concat([pgip_ppo_name_2,pgip_hmo_name_2])

pgip_final_name['PO ID'] = pgip_final_name['PO ID'].astype(str).str.strip() 

pgip_final_name['file_name_prefix'].value_counts()
final_pcp_w_gap_and_refill_onlist['PO_ID'] = final_pcp_w_gap_and_refill_onlist['PO_ID'].astype(str).str.strip() 


final_pcp_w_gap_and_refill_onlist_pgip = pd.merge(final_pcp_w_gap_and_refill_onlist,pgip_final_name, how = 'left' , left_on='PO_ID', right_on= 'PO ID')


#Renaming columns

final_pcp_w_gap_and_refill_onlist_pgip.rename(columns={'Diabetes Medications_Futile Date': "Diabetes Medications_Non-recoverable date",
                                         'RASA_Futile Date': "RASA_Non-recoverable date",
                                         'Statins_Futile Date': "Statins_Non-recoverable date"
                                         }, inplace=True)

final_pcp_w_gap_and_refill_onlist_pgip['file_name_prefix'] = final_pcp_w_gap_and_refill_onlist_pgip['file_name_prefix'].astype(str).str.strip()

final_pcp_w_gap_and_refill_onlist_final = final_pcp_w_gap_and_refill_onlist_pgip[['Member Id',
'Contract ID',
'Member First Name',
'Member Last Name',
'Member Date of Birth',
'Hic_Number',
'Diabetes cohort',
'Diabetes Medications_Non-recoverable date',
'RASA cohort',
'RASA_Non-recoverable date',
'Statins cohort',
'Statins_Non-recoverable date',
'Member_Preferred_Phone_Number',
'Member_Alt1_Phone_Number',
'Member_Alt2_Phone_Number',
'Member_Alt3_Phone_Number',
'Member_Alt4_Phone_Number',
'Member_Eligibility_Phone_Number',
'Member_Language_Preference',
'Member_Address_1',
'Member_Address_2',
'Member_City',
'Member_State',
'Member_Zip_Code',
'Member_Country_Code',
'NPI',
'OrgTitle',
'PCP NAME',
'PO_ID',
'PU_ID',
'PracticeUnit',
'SubGroup',
'SubPO_ID',

'Diabetes Medications_Index Date',
'Diabetes Medications_New to Therapy',
'Diabetes Medications_Most Recent Rx',
'Diabetes Medications_NDC',
'Diabetes Medications_Most Recent Fill Date',
'Diabetes Medications_Next Fill Due Date',
'Diabetes Medications_Pharmacy NPI',
'Diabetes Medications_Pharmacy Name',
'Diabetes Medications_Pharmacy Phone',
'Diabetes Medications_Prescriber NPI',
'Diabetes Medications_Days Supply to Adherent',
'Diabetes Medications_PDC (YTD)',
'Diabetes Medications_Current Fill Status',
'Diabetes Med Refills Remaining',
'Diabetes Med Day Supply',
'Diabetes Medications_Fill Count (YTD)',

'RASA_Index Date',
'RASA_New to Therapy',
'RASA_Most Recent Rx',
'RASA_NDC',
'RASA_Most Recent Fill Date',
'RASA_Next Fill Due Date',
'RASA_Pharmacy NPI',
'RASA_Pharmacy Name',
'RASA_Pharmacy Phone',
'RASA_Prescriber NPI',
'RASA_Days Supply to Adherent',
'RASA_PDC (YTD)',
'RASA_Current Fill Status',
'RASA Med Refills Remaining',
'RASA Med Day Supply',
'RASA_Fill Count (YTD)',

'Statins_Index Date',
'Statins_New to Therapy',
'Statins_Most Recent Rx',
'Statins_NDC',
'Statins_Most Recent Fill Date',
'Statins_Next Fill Due Date',
'Statins_Pharmacy NPI',
'Statins_Pharmacy Name',
'Statins_Pharmacy Phone',
'Statins_Prescriber NPI',
'Statins_Days Supply to Adherent',
'Statins_PDC (YTD)',
'Statins_Current Fill Status',
'Statins Med Refills Remaining',
'Statins Med Day Supply',
'Statins_Fill Count (YTD)',
'Diab Gap Days remaining',
'RAS Gap Days remaining',
'Statin Gap Days remaining',
'file_name_prefix',
'Meaningful_risk_Ind'
]]

#Create folder by the date of File and export in cvs


final_pcp_w_gap_and_refill_onlist_final_only_AHA = final_pcp_w_gap_and_refill_onlist_final[final_pcp_w_gap_and_refill_onlist_final['SubGroup'] == 'Accountable Healthcare Advantage']

final_pcp_w_gap_and_refill_onlist_final_without_AHA = final_pcp_w_gap_and_refill_onlist_final[final_pcp_w_gap_and_refill_onlist_final['SubGroup'] != 'Accountable Healthcare Advantage']

directory_path = rf'W:\STARS_2023\Stars Team\Akshay\Outbound_PCP_Patient_list\{extracted_char}'

# Check if the directory exists
if os.path.exists(directory_path):
    # Remove all files and subdirectories in the existing directory
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)
        if os.path.isfile(item_path):
            os.remove(item_path)
        elif os.path.isdir(item_path):
            shutil.rmtree(item_path)
else:
    # Create the directory if it doesn't exist
    os.makedirs(directory_path)


grouped = final_pcp_w_gap_and_refill_onlist_final_without_AHA.groupby('PO_ID')

for group_name, group_df in grouped:
    file_name = f"{directory_path}\\{group_df['file_name_prefix'].iloc[0]}_{group_name}_{extracted_char}_EOY_MedAdh_Target_Lists.csv"
    group_df.to_csv(file_name, index=False)
    print(f"{group_name} {group_df['OrgTitle'].iloc[0]} file create")


final_pcp_w_gap_and_refill_onlist_final_only_AHA.to_csv(f"{directory_path}\\RBCE016~Accountable Healthcare Advantage_IRBCE0000016_{extracted_char}_EOY_MedAdh_Target_Lists.csv")


################################# Exporting optum data to csv for tracking ##########################


def find_latest_optum_files(root_folder):
    matching_files = []

    for root, dirs, files in os.walk(root_folder):
        for file in files:
            if file.endswith(".xlsx") and "BCBSMISHS_pdc_extract_enh" in file and "~$" not in file:
                file_path = os.path.join(root, file)
                
                # Get the creation time of the file
                created_time = os.path.getctime(file_path)

                # Add a tuple with file path and creation time to the matching_files list
                matching_files.append((file_path, created_time))

    # Sort the list of tuples based on creation time in descending order
    matching_files.sort(key=lambda x: x[1], reverse=True)

    # Return the file path of the latest file
    return matching_files[0][0] if matching_files else None



optum_root = r'W:\STARS_2023\Optum\Advanced Analytics Reporting'
optum_latest_file =  find_latest_optum_files(optum_root)
#print(optum_latest_files)

print(f' Loading....{optum_latest_file}')

optum_file_date = re.search(r'\\(\d+\.\d+\.\d+)\\',optum_latest_file).group(1)

optum_enh_extract = pd.read_excel(optum_latest_file)
print("Optum file read")
optum_enh_extract.to_csv(rf"W:\STARS_2023\Stars Team\Akshay\OW transitions\BCBSM Rx Stars OW Handover materials_with Python repository\4. Core analysis (Python notebooks)\20231010 BCBSM Rx Stars_Python analysis\Claims\BCBSMISHS_pdc_extract_enh_received_{optum_file_date}.csv", index = False)
print("CSV optum claims exported")


############################################## Send a Confirmation Email #########################################

print("Sending email")

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
#mail.To = 'akshay.patel@emergentholdingsinc.com'
mail.To = 'Akshay.Patel@bcbsm.com ; TConnelly@bcbsm.com ; jgembarski@bcbsm.com'
mail.Subject = 'Automated Email - PDC data Refresh complete'
mail.Body = 'Pharmacy refresh complete '
mail.HTMLBody = '<h3>Rx_Ante patient list refresh successful. Please do a random check before sharing</h3>' #this field is optional

# To attach a file to the email (optional):
#attachment  = "Path to the attac
#mail.Attachments.Add(attachment)

mail.Send()

print("Email sent")
print("All process completed successfully")




