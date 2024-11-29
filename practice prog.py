import datetime
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import mysql.connector


current_datetime = datetime.datetime.now()

# Extract the date, time, and day
current_date = current_datetime.date()
current_time = current_datetime.time()
current_day = current_datetime.strftime("%A")
whichclass=""
# Print the results
print("Current Date:", current_date)
print("Current Time:", current_time)
print("Current Day:", current_day)

# Convert the time and date to strings
current_time_str = current_time.strftime("%H:%M:%S")
current_date_str = current_date.strftime("%d-%m-%y")



    # Load the existing workbook
print("Class 4 attendance")
dfclass1=pd.read_excel("G:/python project/Class 4.xlsx")
    # Select the active sheet
imput=dfclass1['28-06-2023']
for i in range(31,41,1):
    summa=input('Roll no '+str(i)+" : ")
    imput[i]=summa

print(imput)
'''Write the DataFrame to the Excel file
df.to_excel(excel_file, index=False)

print('Data written to Excel file:', excel_file)'''


