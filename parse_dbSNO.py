# parse_dbSNO.py - Parses the .html output files from the fetch script and inputs values into the original excel file
# Hamaad Markhiani - June 28, 2016

# imports
import openpyxl			# may have to download via command_line_username$ pip install openpyxl
from bs4 import BeautifulSoup	# may have to download via command_line_username$ pip install beautifulsoup4

# variables
input_excel_name	= "workbook.xlsx"
output_excel_name	= "workbook_out.xlsx"
input_sheet_name	= "Sheet1"
start_column		= "L"
start_row		= "2"
end_column		= "L"
end_row			= "3303"
input_column2		= "M"
output_column		= "N"
count			= 0

# program starts here
# -------------------
# opens the Excel file and loads it into an object
workbook = openpyxl.load_workbook(input_excel_name, data_only=True)

# opens the specific sheet you want to work on
sheet = workbook.get_sheet_by_name(input_sheet_name)

# function to loop until all cells have been queried
def parse_html(start_cell, end_cell, output_co1, output_co2):
	# beautiful soup 4's object for the html file
	soup = ""
	# assign cell range
	cell_range = sheet[start_cell:end_cell]

	# takes value of each cell in range and makes query
	for row in cell_range:
		for cell in row:
			if cell.value:
				soup = BeautifulSoup(open("fetch_dbSNO/" + str(cell.value) + ".html"), "html.parser") # Open the local database file previously fetched
				local_soup = soup.findAll("td", {"class":"style2", "bgcolor":"#F1F1F1", "align":"center"}) # To filter out only the 'td' HTML tags
				i = 0
				for elements in local_soup:
					if local_soup[i].getText() == "IPR000308":	# Marker used to find beginning of search results
						local_soup = local_soup[i+1:]
						break;
					i += 1

				i = 0
				for elements in local_soup:
					if local_soup[i].getText() == str(sheet[input_column2 + str(cell.row)].value):	# if theres a numerical match
						if sheet[output_column + str(cell.row)].value:	# if there is already a value inserted, start hierarchy, vivo > vitro > -,none
							if "vivo" in local_soup[i+3].getText():
								sheet[output_column + str(cell.row)].value = "vivo"
							elif "vitro" in local_soup[i+3].getText() and "vivo" not in local_soup[i+3].getText():
								if "vivo" not in sheet[output_column + str(cell.row)].value:
									sheet[output_column + str(cell.row)].value = "vitro"
						else:	# if empty, insert whatever value
							if "-" in local_soup[i+3].getText():
								sheet[output_column + str(cell.row)].value = "NONE"
							elif "vivo" in local_soup[i+3].getText():
								sheet[output_column + str(cell.row)].value = "vivo"
							elif "vitro" in local_soup[i+3].getText() and "vivo" not in local_soup[i+3].getText():
								sheet[output_column + str(cell.row)].value = "vitro"


					print local_soup[i].getText() # just for debugging
					i += 1
# run the function
parse_html((start_column+start_row), (end_column+end_row), input_column2, output_column)

workbook.save(output_excel_name) # save file!
