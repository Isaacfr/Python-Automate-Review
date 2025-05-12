from bs4 import BeautifulSoup
import requests
import pandas as pd

def compare_gsa_rate_function(year = 2025, zip_code = 20001, check_month='3'):

   #Utilize URL for web-scrapping
   url = f'https://www.gsa.gov/travel/plan-book/per-diem-rates/per-diem-rates-results?action=perdiems_report&fiscal_year={year}&state=&city=&zip={zip_code}'
   page = requests.get(url)
   soup = BeautifulSoup(page.text, features='html')

   #Create a default conversion map that converts a 3-letter date into a digit in string form
   month_mapping = {
       'jan': '1',
       'feb': '2',
       'mar': '3',
       'apr': '4',
       'may': '5',
       'jun': '6',
       'jul': '7',
       'aug': '8',
       'sep': '9',
       'oct': '10',
       'nov': '11',
       'dec': '12'
   }

   #Testing
   #print(soup)
   #soup.find('table')
   #soup.find_all('table')[1]

   #For this website, find the specific table and class to extract information
   table = soup.find('table', class_ = 'table_perdiem')

   #Testing to make sure identity is correct
   #table = soup.find('table', class_ = 'table_perdiem stripedTable dataTable no-footer dtr-inline')
   #table = soup.find_all('table')[1]
   #print(table)

   #Extract Headers
   world_titles = table.find_all('th')
   world_table_titles = [title.text.strip().replace('\xa0', '') for title in world_titles]

   #Testing values and what kind of data is being extracted
   #print(world_table_titles)
   #months = dict.fromkeys(world_table_titles[2:], None)
   months = world_table_titles[2:]
   new_months = []

   #Some data includes a year as well, this is to extract the 3-letter date
   for month in months:
         month_str = month[-3:].lower()
         new_months.append(month_mapping[month_str])

   #Testing difference
   #print(months)
   #print(new_months)

   #Create a dataframe object to extract data
   df = pd.DataFrame(columns = new_months)

   #Find all the rows and insert them into the dataframe
   #Create a blank dictionary for month data
   column_data = table.find_all('tr')
   months_dict = {}

   #Reiterate to each data and extract data, besides headers. 
   for row in column_data[1:]:
        row_data = row.find_all('td')
        individual_row_data = [data.text.strip() for data in row_data[2:]]
    
        #print(individual_row_data)
        length = len(df)
    
        df.loc[length] = individual_row_data

   row_index = 0
   months_dict = df.iloc[row_index].to_dict()

   #print(df)
   #print(months_dict)
   
   #Extract rate based on month that was used as input
   lodging_rate = months_dict[check_month]
   #print(lodging_rate)
   
   return lodging_rate

#Test calling the function
#compare_gsa_rate_function()