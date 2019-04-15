
import requests, json, openpyxl
from datetime import datetime,timedelta


#scrap.py Program to get rate and availability data from booking engine

print("Program to fetch rate and availability")

todayDate = datetime.now()

#getting and validating input for start date
while True:
	print("Please enter the start date (dd-mm-yy): ")
	inStartDate = input()
	try:
		start_date = datetime.strptime(inStartDate[:2]+'-'+inStartDate[3:5]+'-'+inStartDate[6:8], '%d-%m-%y')
		if start_date > todayDate:
			break
		else:
			print("date should be at least tommorrow")
	except:
		print("please try again,,,")

while True:
	max_date = timedelta(days=366)
	print("Please enter end date (dd-mm-yy): ")
	inEndDate = input()
	try:
		end_Date_User = datetime.strptime(inEndDate[:2]+'-'+inEndDate[3:5]+'-'+inEndDate[6:8], '%d-%m-%y')
		if end_Date_User < start_date:
			print("End date should be greater than start date")
		elif end_Date_User > (start_date+max_date):
			print("End date should not be more than a year")
		else:
			break
	except:
		print("please try again,,,")

print("Please wait ,,, I'm preparing :)") 
#initialize variables to go to Booking Engine


timeneeded = 12
startUrl=[]
endUrl=[]
addduration=timedelta(days=28)
data=[]


#getting json data from Booking Engine and save the file as mydata.json
for i in range(timeneeded):
	end_date=start_date+addduration
	startUrl.append(start_date.strftime("%Y-%m-%d"))
	if end_date > end_Date_User:
		end_date=end_Date_User
		endUrl.append(end_date.strftime("%Y-%m-%d"))
		response=requests.get(f'https://app.thebookingbutton.com/api/v1/properties/monarchhousedirect/rates.json?start_date={startUrl[i]}&end_date={endUrl[i]}')
		data.append(json.loads(response.text))
		break
	else:
		endUrl.append(end_date.strftime("%Y-%m-%d"))
		response=requests.get(f'https://app.thebookingbutton.com/api/v1/properties/monarchhousedirect/rates.json?start_date={startUrl[i]}&end_date={endUrl[i]}')
		data.append(json.loads(response.text))
		start_date=end_date+timedelta(days=1)
	

#saving data to JSON
with open("mydatad.json","w") as datafile:
	json.dump(data, datafile, indent=4)
	datafile.close()

try:
	response.raise_for_status()
	print("Data downloaded")
except:
	print("hmm, something is wrong")


print("preparing ..")
#Loading Json data and preparing to save to an excel file
rawdata = open('mydatad.json')
data = json.load(rawdata)

wb = openpyxl.Workbook()
sheet = wb.active

#setting up headers,merging and printing
title1 = "DATE"
title2 = "AVAILABLE"
title3 = "NET RATE"
sheet.merge_cells('A1:C1')
sheet.merge_cells('F1:H1')
sheet.merge_cells('K1:M1')
sheet['A2'] = title1 
sheet['B2'] = title2 
sheet['C2'] = title3
sheet['F2'] = title1 
sheet['G2'] = title2 
sheet['H2'] = title3
sheet['K2'] = title1 
sheet['L2'] = title2 
sheet['M2'] = title3
sheet.freeze_panes = 'A3'
theColumnWidth = 12
sheet.column_dimensions['B'].width = theColumnWidth
sheet.column_dimensions['G'].width = theColumnWidth
sheet.column_dimensions['L'].width = theColumnWidth

#initiate variable for row loops in excel
C1 = 3
C2 = 3
C3 = 3

#variable for columns
COL = 1
x=1

print("preparing ...")
#Looping data to extract information needed
for x in data:
	
	#print(type(data))
	for y in x:
		
		for z in range(len(y["room_types"])):
			pdate = y["room_types"][z]["room_type_dates"]

			if y["room_types"][z]["name"] == "One Bedroom Apartment":
				#printing header in Excel
				sheet.cell(row=1, column=COL).value = y["room_types"][z]["name"]
				for a in range(len(y["room_types"][z]["room_type_dates"])):
					sheet.cell(row=C1, column=COL).value = pdate[a]["date"][8:10] + " / " + pdate[a]["date"][5:7]
					sheet.cell(row=C1, column=COL+1).value = y["room_types"][z]["room_type_dates"][a]["available"]
					sheet.cell(row=C1, column=COL+2).value = round(int(y["room_types"][z]["room_type_dates"][a]["rate"])/1.15)
					C1 += 1
			
			if y["room_types"][z]["name"] == "Two Bedroom Apartment":
				#printing header in Excel
				sheet.cell(row=1, column=COL+5).value = y["room_types"][z]["name"]
				for a in range(len(y["room_types"][z]["room_type_dates"])):
					sheet.cell(row=C2, column=COL+5).value = pdate[a]["date"][8:10] + " / " + pdate[a]["date"][5:7] 
					sheet.cell(row=C2, column=COL+6).value = y["room_types"][z]["room_type_dates"][a]["available"]
					sheet.cell(row=C2, column=COL+7).value = round(int(y["room_types"][z]["room_type_dates"][a]["rate"])/1.15)
					C2 += 1

			if y["room_types"][z]["name"] == "Superior Three Bed":
				#printing header in Excel
				sheet.cell(row=1, column=COL+10).value = y["room_types"][z]["name"]
				for a in range(len(y["room_types"][z]["room_type_dates"])):
					sheet.cell(row=C3, column=COL+10).value = pdate[a]["date"][8:10] + " / " + pdate[a]["date"][5:7]
					sheet.cell(row=C3, column=COL+11).value = y["room_types"][z]["room_type_dates"][a]["available"]
					sheet.cell(row=C3, column=COL+12).value = round(int(y["room_types"][z]["room_type_dates"][a]["rate"])/1.15)
					C3 += 1


print("Done :)")
	
wb.save("MonarchHouseAvailability.xlsx")



#end of program