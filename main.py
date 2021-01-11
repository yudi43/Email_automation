from psycopg2 import sql
from psycopg2 import Error
import psycopg2
from datetime import datetime
import openpyxl
# Define schema and table names respectively
schema_name = 'public'
table_name = 'DailySiteReportData';
connection = psycopg2.connect(user = "YOUR_DB_USER", password = "YOUR_DB_PASS", host = "localhost", port = "5433", database = "YOUR_DB_NAME")
cursor = connection.cursor()
entries_made_today = [] 
# Get all the columns information
cursor.execute(sql.SQL("SELECT * FROM {}.{}").format(
    sql.Identifier(schema_name),
    sql.Identifier(table_name)
))
colnames = [desc[0] for desc in cursor.description] # This list will contain all the column names.

# Get the index of column named "date"
index_of_date = colnames.index('date')
todays_date = datetime.today().strftime('%Y-%m-%d')

table = cursor.fetchall() 
for entry in table:
    if entry[index_of_date].strftime('%Y-%m-%d') == todays_date:
        entries_made_today.append(entry)

# Use the list "entries_made_today" to create sheets based on number of items and fill all the sheets accordingly!
#####################################################


# FOR NOW WE ARE GOING TO FILL JUST ONE SHEET (sheet 1)
sheet1_data = None
if (len(entries_made_today) > 0):
    sheet1_data = entries_made_today[0]
print("This is the sheets 1 data: ", sheet1_data)
in_file_name = "test.xlsx"
in_path = '/Users/yudi/Desktop/' + in_file_name;
xfile = openpyxl.load_workbook(in_path)
sheet1 = xfile.get_sheet_by_name('Sheet1')
print(colnames.index('date'))
sheet1['D3'] = sheet1_data[colnames.index('sub_contractor_name')]
sheet1['D4'] = sheet1_data[colnames.index('project_statement')]
sheet1['D5'] = sheet1_data[colnames.index('job_number_description')]
sheet1['D8'] = sheet1_data[colnames.index('date')]
sheet1['D9'] = sheet1_data[colnames.index('contractor_name')]
sheet1['D10'] = sheet1_data[colnames.index('working_time')]
sheet1['D11'] = sheet1_data[colnames.index('working_hours')]
sheet1['D12'] = sheet1_data[colnames.index('shift_work')]
sheet1['D13'] = sheet1_data[colnames.index('weather_conditions')]
sheet1['E17'] = sheet1_data[colnames.index('number_engineers_site_manager')]
sheet1['E18'] = sheet1_data[colnames.index('number_foremen_supervisors')]
sheet1['E19'] = sheet1_data[colnames.index('number_surveyors')]
sheet1['E20'] = sheet1_data[colnames.index('number_quantity_surveyors')]
sheet1['E21'] = sheet1_data[colnames.index('number_administrative_personnel')]
sheet1['E22'] = sheet1_data[colnames.index('number_concrete_workers')]
sheet1['E23'] = sheet1_data[colnames.index('number_cladding_working_roofers')]
sheet1['E24'] = sheet1_data[colnames.index('number_steel_fixers_bar_handlers')]
sheet1['E25'] = sheet1_data[colnames.index('number_steel_workers_welders')]
sheet1['E26'] = sheet1_data[colnames.index('number_electricians_operators_drivers_security_and_other_skille')]
sheet1['E27'] = sheet1_data[colnames.index('number_painters')]
sheet1['E28'] = sheet1_data[colnames.index('number_masons_plasters')]
sheet1['E29'] = sheet1_data[colnames.index('number_other_skilled_workers')]
sheet1['E30'] = sheet1_data[colnames.index('number_unskilled_workers')]
sheet1['E31'] = sheet1_data[colnames.index('number_visitors')]

sheet1['F17'] = sheet1_data[colnames.index('comments_engineers_site_manager')]
sheet1['F18'] = sheet1_data[colnames.index('comments_foremen_supervisors')]
sheet1['F19'] = sheet1_data[colnames.index('comments_surveyors')]
sheet1['F20'] = sheet1_data[colnames.index('comments_quantity_surveyors')]
sheet1['F21'] = sheet1_data[colnames.index('comments_administrative_personnel')]
sheet1['F22'] = sheet1_data[colnames.index('comments_concrete_workers')]
sheet1['F23'] = sheet1_data[colnames.index('comments_cladding_working_roofers')]
sheet1['F24'] = sheet1_data[colnames.index('comments_steel_fixers_bar_handlers')]
sheet1['F25'] = sheet1_data[colnames.index('comments_steel_workers_welders')]
sheet1['F26'] = sheet1_data[colnames.index('comments_electricians_operators_drivers_security_and_other_skil')]
sheet1['F27'] = sheet1_data[colnames.index('comments_painters')]
sheet1['F28'] = sheet1_data[colnames.index('comments_masons_plasters')]
sheet1['F29'] = sheet1_data[colnames.index('comments_other_skilled_workers')]
sheet1['F30'] = sheet1_data[colnames.index('comments_unskilled_workers')]
sheet1['F31'] = sheet1_data[colnames.index('comments_visitors')]

sheet1['E32'] = sheet1_data[colnames.index('total_personnel')]
sheet1['D35'] = sheet1_data[colnames.index('rental_description')]
sheet1['D36'] = sheet1_data[colnames.index('rental_quantity')]
sheet1['D39'] = sheet1_data[colnames.index('covid19_personnel_count_affected')]
sheet1['D40'] = sheet1_data[colnames.index('covid19_more_details')]
sheet1['C44'] = sheet1_data[colnames.index('executed_works')]
sheet1['C50'] = sheet1_data[colnames.index('unusual_events_accidents_work_interruptions')]
sheet1['C56'] = sheet1_data[colnames.index('materials_received')]
sheet1['C62'] = sheet1_data[colnames.index('tools_equipment_used')]
sheet1['D70'] = sheet1_data[colnames.index('responsible_contractors_representative_on_site')]
sheet1['D71'] = sheet1_data[colnames.index('date')]
sheet1['D73'] = sheet1_data[colnames.index('harris_pye_site_manager')]
sheet1['D74'] = sheet1_data[colnames.index('date')]
# sheet1['D76'] = sheet1_data[colnames.index('signed')]

sheet1['D76'] = sheet1_data[colnames.index('user_fullname')]
sheet1['D77'] = sheet1_data[colnames.index('date_entered')]




xfile.save('text2.xlsx')


###################################################

