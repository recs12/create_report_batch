#!/usr/bin/env python
# coding: utf-8

# In[164]:


import pandas as pd
import glob

def read(excel_file):
    df = pd.read_excel(excel_file, dtype={"ID":str, "Revision":str,})
    return df

# Append all dataframes.
batch = pd.DataFrame()
for f in glob.glob("*.xlsx"):
    df =read(f)
    batch = batch.append(df,ignore_index=True, sort=False)

# Remove 'PaperVersion', 'PT_Garbage', 'PT_Obsolete' in column 'Release Status'.
batch = batch[(batch['Release Status']!='PaperVersion') & (batch['Release Status']!='PT_Garbage')& (batch['Release Status']!='PT_Obsolete')]

# Sorting by ID and Revision.
batch.sort_values(by=['ID', 'Revision'], ascending=True)

# Keep those three columns.
batch = batch[['ID','Revision','Current Name']]

# Keep the latest revision.
batch = batch.drop_duplicates(subset='ID', keep='last')

# Generate excel report.
batch.to_excel("batch.xlsx", index=False)

