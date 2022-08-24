import xlsxwriter
import requests

#Receiving data from HTTP GET request
data = requests.get('https://restcountries.com/v3.1/all')

#Creating workbook and worksheet
file = xlsxwriter.Workbook('Countries_List.xlsx')
worksheet = file.add_worksheet()

#Defining the formatting that will be used while writing data into cells
merge_format = file.add_format({'align': 'center'})
title_format = file.add_format({'bold':True, 'font_size': 16, 'font_color':'#4F4F4F', 'align': 'center'})
column_name_format = file.add_format({'bold': True, 'font_size':12, 'font_color': '#808080'})
number_format = file.add_format({'num_format': '#,###.00'})

#Defining column width for better readability
worksheet.set_column(0,3,20)

#Writing the first titular cells(with the appropriate formatting)
worksheet.merge_range(0,0,0,3,'Countries List', title_format)
worksheet.write(1,0,'Name',column_name_format)
worksheet.write(1,1,'Capital',column_name_format)
worksheet.write(1,2,'Area',column_name_format)
worksheet.write(1,3,'Currencies',column_name_format)

#Initial row for the actual data to be written
row = 2

#Writing data in a row for each country, using the appropriate formatting and inserting "-" if the corresponding field doesn't exist in the data
for country in data.json():
    worksheet.write(row,0,country['name']['common'])
    worksheet.write(row,1, country['capital'][0] if 'capital' in country else '-')
    worksheet.write(row,2,country['area'], number_format)
    worksheet.write(row,3,[i for i in country['currencies'].keys()][0] if 'currencies' in country else '-')
    row +=1

#Finish writing the file and saving it
file.close()