"""
The functions here are dependent on one another and should be run in the order presented.
The only function input required is the path to the folder containing all the excel files for the function
energy_data_file_reference_function. The functions table_transformer and sheets_to_dataframes need to be
run so the transformed_energy_data class can be created.
""" 
folder_path = "C:/Users/danii/OneDrive/Desktop/Energy Data Files"
# enter the folder path for the energy sector data
def energy_data_file_reference_function(folder_path):
    """
    Given a folder path, creates a dictionary pairing the the filename, the title of the workbook in the 
    contents sheet, the paths, sheetnames, and number of sheets in total.
    """
    # import packages
    import os
    import numpy as np
    import pandas as pd
    
    # get list of all .xlsx files in our data folder
    energy_data_file_names = os.listdir(folder_path)
    
    # list that will store all file paths in energy data folder
    energy_data_file_paths = []

    # loop to create paths for each file and populate above list
    for i in range(len(energy_data_file_names)): 
        path = folder_path + '/' + energy_data_file_names[i]
        energy_data_file_paths.append(path)
    
    full_workbook_titles = []

    # loop to get workbook title using paths
    for i in range(len(energy_data_file_paths)):
        workbook = pd.read_excel(energy_data_file_paths[i])   
        full_workbook_titles.append(workbook.iloc[5,0]) 
        # note: when opening a workbook (above) using pandas, 
        # it omits the first two cells that include the name of the agency and the data source
        # so despite being in cell A7 in excel the title of the workbook is in row 5, column 0 in pandas 
        
    # get number of sheets in each workbook in case we need to write a loop later    
    num_sheets = []

    for i in range(len(energy_data_file_paths)):
        df = pd.ExcelFile(energy_data_file_paths[i])
        sheet_count = len(df.sheet_names)
        num_sheets.append(sheet_count)
        
    # get list of lists of sheet names in the 
    sheet_names = []

    for i in range(len(energy_data_file_paths)):
        df = pd.ExcelFile(energy_data_file_paths[i])
        sheet_names.append(df.sheet_names)
    
    # create a dictionary pairing the the filename, the title of the workbook in the contents sheet,
    # the paths, sheetnames, and number of sheets in total
    energy_data_file_reference = dict(zip(energy_data_file_names, zip(full_workbook_titles, 
                                                                    energy_data_file_paths, 
                                                                    num_sheets, 
                                                                    sheet_names)))
    return energy_data_file_reference

######################################################################################################

def table_transformer(workbook_path, sheet_name):
    """
    Takes the workbook path and sheetname from United States Energy Information Administration
    energy data and transforms into a useable format.
    """
    # import packages
    import numpy as np
    import pandas as pd
    # load workbook
    workbook_loaded = pd.ExcelFile(workbook_path)
    # load individual worksheet
    df = pd.read_excel(workbook_loaded, sheet_name)
    # drop rows with all NaNs
    df.dropna(axis=0, inplace=True, how="all") 
    # drop rows with less than 2 values, in this case, rows with just the years and NaNs
    df.dropna(axis=1, inplace=True, thresh=2)
    # transform table so states are column names
    df = df.T
    df.columns = df.iloc[0]
    # rename 'State' column to 'Year'
    df.rename(columns = {"State": "Year"}, inplace = True)
    # drop duplicate row
    df = df.iloc[1:]
    # set 'Year' Column to int
    df['Year'] = df['Year'].astype('int')
    # reset index 
    df = df.reset_index(drop=True)
    # list of states
    states = {
            'AK': 'Alaska', 'AL': 'Alabama', 'AR': 'Arkansas','AZ': 'Arizona',
            'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DC': 'District of Columbia',
            'DE': 'Delaware', 'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 
            'IA': 'Iowa', 'ID': 'Idaho', 'IL': 'Illinois', 'IN': 'Indiana',
            'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana', 'MA': 'Massachusetts',
            'MD': 'Maryland', 'ME': 'Maine', 'MI': 'Michigan', 'MN': 'Minnesota',
            'MO': 'Missouri', 'MS': 'Mississippi', 'MT': 'Montana', 'NC': 'North Carolina',
            'ND': 'North Dakota', 'NE': 'Nebraska', 'NH': 'New Hampshire', 'NJ': 'New Jersey',
            'NM': 'New Mexico', 'NV': 'Nevada', 'NY': 'New York', 'OH': 'Ohio',
            'OK': 'Oklahoma', 'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island',
            'SC': 'South Carolina', 'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas',
            'UT': 'Utah', 'VA': 'Virginia', 'VT': 'Vermont', 'WA': 'Washington',
            'WI': 'Wisconsin', 'WV': 'West Virginia', 'WY': 'Wyoming', 'US': 'US Total'
        }
    # change state abbreviations to full names
    df.rename(columns = states, inplace = True)
    # return dataframe
    return df

######################################################################################################

def sheets_to_dataframes(folder_path, excel_file):
    """
    Take a folder path and excel file name in that folder to return multiple dataframes equivelent to the 
    relevent (non-contents) sheets.
    Note: this function depends on energy_data_file_reference_function and table_transformer functions.
    """
    # get path to file
    energy_data_file_reference = energy_data_file_reference_function(folder_path)
    workbook_path = energy_data_file_reference[excel_file][1]
    # create empty list to which output tables will be assigned ans stored
    lst_tables = []
    for i in range(len(energy_data_file_reference[excel_file][3])):
        if i == 0: # skip 'Contents' sheet in files
            continue
        else:
            sheet_name = energy_data_file_reference[excel_file][3][i]
            lst_tables.append(table_transformer(workbook_path, sheet_name))
    return lst_tables

######################################################################################################

class transformed_energy_data():
    """
    Transformed energy data for analysis.
    Note: this function depends on the sheets_to_dataframes function.
    """
    import numpy as np
    # Generic 'Year' column for later time series and plotting
    year = np.arange(1960,2021)

    ### Production ### 
    # Primary Energy Production Estimates, Fossil Fuels and Nuclear Energy
    coal_prod, natgas_prod, crude_prod, nuclear_prod = sheets_to_dataframes(folder_path,'prod_btu_ff_nu.xlsx')

    # Primary Energy Production Estimates, Renewable and Total Energy
    biofuel_prod, wood_waste_prod, other_renew_prod, tot_renew_prod, tot_ener_prod = sheets_to_dataframes(folder_path,'prod_btu_re_te.xlsx')
    # Other Renewables are non_combustible like geo, hydro, solar, and wind

    ### Consumption ### 
    # Primary Energy Consumption Estimates by Source
    coal_consump, natgas_consump, petro_consump, nuclear_consump, tot_renew_consump = sheets_to_dataframes(folder_path,'use_energy_source.xlsx')
    # Renewable Energy Consumption Estimates by Source
    biomass_consump, geother_consump, hydro_consump, solar_consump, wind_consump = sheets_to_dataframes(folder_path,'use_renew_sector.xlsx')

    ### Expenditure ###
    # Total Energy Price and Expenditure Estimates
    tot_expend_prices, tot_expend, tot_expend_percap, tot_expend_pergdp = sheets_to_dataframes(folder_path,'pr_ex_tot.xlsx')
    # Motor Gasoline Price and Expenditure Estimates
    motogas_prices, motogas_expend, motogas_expend_percap = sheets_to_dataframes(folder_path,'pr_ex_mg.xlsx')
    # Petroleum and Natural Gas Price and Expenditure Estimates
    petro_prices, petro_expend, natgas_prices, natgas_expend = sheets_to_dataframes(folder_path,'pr_ex_pa_ng.xlsx')
    # Coal and Electricity Retail Sales Price and Expenditure Estimates
    coal_prices, coal_expend, elec_prices, elec_expend = sheets_to_dataframes(folder_path,'pr_ex_cl_es.xlsx')

    ### By Sector ###
    # Total Energy Consumption Estimates by End-Use Sector
    resid_consump, commer_consump, indust_consump, transp_consump, tot_sec_consump = sheets_to_dataframes(folder_path,'use_tot_sector.xlsx')
    # Total Energy Consumption Estimates per Capita by End-Use Sector
    resid_consump_percap, commer_consump_percap, indust_consump_percap, transp_consump_percap, tot_sec_consump_percap = sheets_to_dataframes(folder_path,'use_tot_capita.xlsx')
    # Total Energy Price Estimates by End-Use Sector
    resid_prices, commer_prices, indust_prices, transp_prices, tot_sec_prices = sheets_to_dataframes(folder_path,'pr_avg_tot.xlsx')
    # Total Energy Expenditure Estimates by End-Use Sector
    resid_expend, commer_expend, indust_expend, transp_expend, tot_sec_expend = sheets_to_dataframes(folder_path,'expend_tot.xlsx')
    # Electricity Retail Sales, Total and Residential, Total and per Capita
    tot_ret_elec_sales, tot_ret_elec_sales_percap, resid_elec_sales, resid_elec_sales_percap = sheets_to_dataframes(folder_path,'use_es_capita.xlsx')

    ### Economic Relevant Data ###
    # Total Energy Consumption Estimates, Real Gross Domestic Product (GDP), 
    # Energy Consumption Estimates per Real Dollar of GDP
    tot_consump, realgdp2, ener_consump_per_realgdp = sheets_to_dataframes(folder_path,'use_tot_realgdp.xlsx')
    
    # dataframe that has the figures for Total Primary Energy Production
    from functools import reduce
    tot_prime_prod = reduce(lambda a, b: a.add(b, fill_value=0), [coal_prod.copy(), natgas_prod, crude_prod, nuclear_prod])
    tot_prime_prod['Year'] = year
    tot_prime_prod.columns.name = 'Total Primary Energy Production, Billion Btu'
    tot_prime_prod
    
    # Get Total Primary Energy Consumption so we can compare it with Renewable Energy Production
    tot_prime_consump = reduce(lambda a, b: a.add(b, fill_value=0), [coal_consump.copy(), natgas_consump, petro_consump, nuclear_consump])
    tot_prime_consump['Year'] = year
    tot_prime_consump.columns.name = 'Total Primary Energy Consumption, Billion Btu'
    tot_prime_consump
