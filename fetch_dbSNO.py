# fetch_dbSNO.py - Queries the 'DataBase of Cysteine S-NitrOsylation' and saves the .html output in a local directory
# Hamaad Markhiani - June 28, 2016

# imports
import openpyxl	# may have to download via command_line_user_here$ pip install openpyxl
import urllib2	# usually available

# variables
input_excel_name	= "workbook.xlsx"
input_sheet_name	= "Sheet1"
dbSNO_url_template	= "http://140.138.144.145/~dbSNO/search_result.php?search_type=db_id&swiss_id="
start_column		= "L"
start_row		= "2"
end_column		= "L"
end_row			= "3303"
count			= 0

# program starts here
# -------------------
# opens the Excel file and loads it into an object
workbook = openpyxl.load_workbook(input_excel_name, data_only=True)

# opens the specific sheet you want to work on
sheet = workbook.get_sheet_by_name(input_sheet_name)

# function to loop until all cells have been queried
def query_dbSNO(start_cell, end_cell):

	# assign cell range
	cell_range = sheet[start_cell:end_cell]

	# takes value of each cell in range and makes query
	for row in cell_range:
		for cell in row:
			if cell.value:
				full_url = dbSNO_url_template + str(cell.value) # makes real URL
				content = urllib2.urlopen(full_url) # sends http request
				output = open(str(cell.value) + ".html", "w") # opens output file
				output.write(content.read()) # writes output to output file
				output.close() # closes output file
				global count
				count += 1
				print count
# run the function
query_dbSNO((start_column+start_row), (end_column+end_row))


