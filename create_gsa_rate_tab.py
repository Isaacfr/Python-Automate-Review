import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
import filenames, compare_gsa_rate

def combine_month_and_zip(a, b):
    return a + b

def create_gsa_column(row):
    #Combine month and zip code column to search for key
    #If values do not exist, call a function to web-scrap the value and add it to the dictionary and return the value
    combined_key = combine_month_and_zip(row['Month'], row['Hotel Postal Code'])
    if combined_key in result_dict and result_dict[combined_key] is not None:
        return result_dict[combined_key]
    else:
        val = compare_gsa_rate.compare_gsa_rate_function(zip_code = row['Hotel Postal Code'], check_month = row['Month'])
        result_dict[combined_key] = val
        return val

def main():
    # Read the Excel file into a DataFrame
    file_path = 'your_file.xlsx'
    df = pd.read_excel(filenames.source_file, filenames.raw_sheet)  # Specify the sheet name if necessary

    #Select specific columns
    columns_to_extract = ['Booking ID', 'Billed End', 'Average Nightly Rate w/out Taxes and Fees', 'Hotel Postal Code', 'Travelers']
    extracted_df = df[columns_to_extract]

    #Format the following tabs to correct format for web-scrapping and Dataframe
    extracted_df['Billed End'] = pd.to_datetime(df['Billed End'], format='%m%d%Y')
    extracted_df['Month'] = extracted_df['Billed End'].dt.month.map(str)
    extracted_df['Hotel Postal Code'] = extracted_df['Hotel Postal Code'].map(str).fillna('').str.split('.').str[0]

    #Create a unique list of values to only have to repeat web-scrap unique values.
    combined_list = (combine_month_and_zip(extracted_df['Month'], extracted_df['Hotel Postal Code'])).tolist()
    unique_values = set(combined_list)

    #Initialize a dictionary that is set to none
    global result_dict
    result_dict = {key: None for key in unique_values}

    #Create a new tab that finds the gsa_rate for each row.
    extracted_df['GSA Rate'] = extracted_df.apply(create_gsa_column, axis=1)

    #print(result_dict)
    #print(extracted_df)

    #Save the file and do not overwrite the current tabs
    with pd.ExcelWriter(filenames.source_file, engine='openpyxl', mode='a') as writer:
        extracted_df.to_excel(writer, sheet_name='GSA Rate', index=False)