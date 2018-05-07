import os
import re
import sys
import pandas as pd
import xlsxwriter

# Define some data types for columns where we don't want pandas to
# guess the datatype.
DataTypes = {
    'Home_Phone': object,
    'Mobile': object,
    'DepartmentCode': object
         }

# Get list of all files with the .csv extension
# Executable will take up to 1 argument, which is the location of the input
# files. If this is left blank, the default value is the current working
# directory.

# Use regular expressions to search for files with the .csv extension only
suffix = re.compile(".csv", re.I)

# Set input directory as the first argument
if len(sys.argv) > 1:
    workdir = sys.argv[1]
    input_paths = [f for f in os.listdir(workdir) if suffix.search(f)]
else:
    workdir = os.getcwd() # Input will be from working directory if no arguments are specified
    input_paths = [f for f in os.listdir(workdir) if suffix.search(f)]

# Define writer for excel destination
writer = pd.ExcelWriter(os.path.join(workdir, 'Calling Lists.xlsx'), engine='xlsxwriter')
    
# Loop through all csv files in working directory and
# write them into excel workbook as separate tabs
for filename in input_paths:
    print 'Writing ' + filename + ' to file ...'
    df = pd.read_csv(os.path.join(workdir, filename), dtype=DataTypes)
    curSheet = filename.split(".")[0]
    df.to_excel(writer, index=False, sheet_name=curSheet)

    # Format the output
    workbook = writer.book
    worksheet = writer.sheets[curSheet]
    worksheet.set_column('B:B', 32)
    worksheet.set_column('C:E', 13)
    worksheet.set_column('F:G', 28)
    worksheet.set_column('H:H', 19)
    worksheet.set_column('I:I', 11.14)
    worksheet.set_column('J:J', 14.86)
    worksheet.set_column('P:P', 15.71)

writer.save()
