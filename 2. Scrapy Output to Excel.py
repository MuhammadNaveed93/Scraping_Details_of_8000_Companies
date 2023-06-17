# -*- coding: utf-8 -*-
"""
Created on Wed Jun 14 20:48:30 2023

@author: muhammad.naveed
"""


import pandas as pd
import Job02_function

df = pd.read_csv("D:\\Naveed\\Coursera\\1. Naveed\\1. Python course\\3. Self Learning\\Programs\\15. Scrapy Projects\\seleniumscrapy\\project_output_final.csv")

data = []

for index, row in df.iterrows():
    fax_list = str(row['fax']).split(',') if not pd.isnull(row['fax']) else ['']
    telephone_list = str(row['telephone']).split(',') if not pd.isnull(row['telephone']) else ['']
    telephone_fax_list = str(row['telephone_fax']).split(',') if not pd.isnull(row['telephone_fax']) else ['']
    email_list = str(row['email']).split(',') if not pd.isnull(row['email']) else ['']
    website_list = str(row['website']).split(',') if not pd.isnull(row['website']) else ['']
    mobile_list = str(row['mobile']).split(',') if not pd.isnull(row['mobile']) else ['']

    max_length = max(len(fax_list), len(telephone_list), len(telephone_fax_list), len(email_list), len(website_list), len(mobile_list))

    other_values = [row['Company_name'], row['Company_address'], row['Town_city'], row['Province'], row['Region_state'], row['Country'], row['Latitude'], row['Longitude'], row['links']]

    for i in range(max_length):
        row_data = other_values + [fax_list[i] if i < len(fax_list) else '',
                                   telephone_list[i] if i < len(telephone_list) else '',
                                   telephone_fax_list[i] if i < len(telephone_fax_list) else '',
                                   email_list[i] if i < len(email_list) else '',
                                   website_list[i] if i < len(website_list) else '',
                                   mobile_list[i] if i < len(mobile_list) else '']
        data.append(row_data)

new_df = pd.DataFrame(data, columns=['Company_name', 'Company_address', 'Town_city', 'Province', 'Region_state', 'Country', 'Latitude', 'Longitude', 'links', 'fax', 'telephone', 'telephone_fax', 'email', 'website', 'mobile'])

#print(new_df[:5])

Job02_function.Convert_to_Excel(new_df, "Project_output")

Job02_function.Merge_cells("Project_output")