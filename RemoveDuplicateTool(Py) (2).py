#!/usr/bin/env python
# coding: utf-8

# In[16]:


import pandas as pd

# Read the Excel file
file_path = r'C:\Users\RM\Documents\JAN FILE.xlsx'

# Read all sheet names from the Excel file
xl = pd.ExcelFile(file_path)
sheet_names = xl.sheet_names

# Initialize an empty dictionary to store grouped DataFrames
grouped_dfs = {}

# Process each sheet
for sheet_name in sheet_names:
    # Read data from the current sheet
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Convert columns to lowercase
    for col in df.columns:
        df[col] = df[col].astype(str).str.lower()

    # Assuming different column names for each sheet
    if sheet_name == 'SELLERCAT':
        group_col = 'Seller Name'
        agg_col = 'Category L2'
        
    elif sheet_name == 'SELLERLIST':
        group_col = 'Seller Name'
        agg_col = 'Listing Name'

    else:
        # Add additional elif blocks for more sheets if needed
        continue

    # Group by the specified column and aggregate another column into a list
    grouped_df = df.groupby(group_col)[agg_col].agg(', '.join).reset_index()

    # Store the grouped DataFrame in the dictionary with the original case sheet name
    grouped_dfs[sheet_name] = grouped_df

# Write each grouped DataFrame to a separate sheet in the Excel file
with pd.ExcelWriter(r'C:\Users\RM\Documents\RESULTT.xlsx', engine='xlsxwriter') as writer:
    for sheet_name, grouped_df in grouped_dfs.items():
        grouped_df.to_excel(writer, sheet_name=sheet_name, index=False)

print('DONE')


# In[ ]:




