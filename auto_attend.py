import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import date


# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheet = client.open("PyLO Attendance").get_worksheet(1)


#Returns the common row. Returns 0 if there is no common row
def compare_lists(lst1, lst2):
	
	rowdict = {row:1 for row in lst1}
	for row in lst2:
		if row in rowdict:
			return row
	return 0

def remove_left_zeros(value):
	for num in value:
		if num == '0':
			value = value[1:]
		else:
			return value
	return value

#Forces strings to be all lowercase except for the first letter
def adjust_text(txt):
	txt = list(txt.lower().strip())
	txt[0] = txt[0].upper()
	for index in range(len(txt) - 1):
		if txt[index] in ("-", " "):
			txt[index + 1] = txt[index + 1].upper()
	return ''.join(txt)


def pull_EC_students(sumcol):
	fname_list = sheet.col_values(1)
	lname_list = sheet.col_values(2)
	sumlist = sheet.col_values(sumcol)
	totallist = zip(fname_list, lname_list, sumlist)

	print("List of students with 4 or more projects completed:")
	completedlist = list(filter(lambda a: a[2].isdigit() and int(a[2]) >= 8, totallist))

	print([f'{elem[0]} {elem[1]}' for elem in completedlist])

def call_attendance(datecol, sumcol):

	firstname = input("What is your first name? ")
	firstname = adjust_text(firstname)

	try:
		row_list_f = [cell.row for cell in sheet.findall(firstname)]

		if len(row_list_f) > 1:
			
			lastname = input("What is your last name? ")
			lastname = adjust_text(lastname)

			try:
				row_list_l = [cell.row for cell in sheet.findall(lastname)]

				finalrow = compare_lists(row_list_f, row_list_l)

				#No first name last name combination
				if not finalrow:
					print("That name combination is either not registered or registered improperly on our attendance sheet.\nPlease let Brian know.\n")
					return

			#Nonexistant last name
			except:
				print("That last name is either not registered or registered improperly on our attendance sheet.\nPlease let Brian know.\n")
				return

		else: 
			finalrow = row_list_f[0]

	#Nonexistant first name
	except:
		print("That first name is either not registered or registered improperly on our attendance sheet.\nPlease let Brian know.\n")
		return

	lastname = sheet.cell(finalrow, 2).value

	if sheet.cell(finalrow, datecol).value:

		print(f'You\'ve already filled attendance for this week, {firstname} {lastname}.\n')

	else:
		sheet.update_cell(finalrow, datecol, '1')

		personal_sum = sheet.cell(finalrow, sumcol).value
		if personal_sum == '':
			personal_sum = '0'

		sheet.update_cell(finalrow, sumcol, str(int(personal_sum) + 1))

		print(f'Attendance recorded for {firstname} {lastname}.\n')


def main():

	today = date.today().strftime("%m/%d/%Y")
	final_date = '/'.join([remove_left_zeros(value) for value in today.split("/")])

	try:
		datecol = sheet.find(final_date).col

	except:
		print("You are trying to access the sheet on an invalid date.\n")
		return

	sumcol = sheet.find("SUM").col

	while(True):
		call_attendance(datecol, sumcol)

if __name__ == "__main__":
	main()


