# Create PS Course File Structure Project
# Project to create folder, read excel, and powerpoint template
#======================================================

import pandas as pd

# read .xlsx file
df = pd.read_excel('Module Overview Course.xlsx')

column_name = 'Course'
df[column_name] = df[column_name].astype(str)

print(df[column_name])

# Create Course Folder


# Create Module Folder

# Change Presentation Title in Empty .pptx

# Copy and Rename Empty pptx to Module Folder

# Conditional check Last Module in file