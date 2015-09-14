import win32com.client

excel = win32com.client.Dispatch('Excel.Application')

# Asks for info to be written to excel file
mouse_breed = int(
    raw_input("for CBS: 1,G6PD:2, Gpx1KO: 3, TgGpx1: 4, Gpx3ko: 5, Gpx1/3ko: 6, Hg6pd: 7, BMPR2: 8, GPP29: 9"))
first_Male = int(raw_input('enter first male\'s number'))
last_Male = int(raw_input('enter last male\'s number'))
first_Female = int(raw_input('enter first female\'s number'))
last_Female = int(raw_input('enter last female\' number'))
DOB = raw_input('enter date of birth')
Nmales = int(raw_input('# of males in litter?'))
Nfemales = int(raw_input('# of females in litter?'))
sireID = raw_input("what was the sire's ID #?")
sireGT = raw_input("what was the sire's genotype?")
damID = raw_input("what were the dams' ID #'s?")
damGT = raw_input("what were the dams' genotypes?")

BMPR2_males = range(first_Male, last_Male + 1)
Male_animals = range(first_Male, last_Male + 1, 2)
Female_animals = range(first_Female, last_Female + 1, 2)
breed = {1: 'CBS', 2: 'G6PD', 3: 'Gpx1KO', 4: 'TgGpx1', 5: 'Gpx3ko', 6: 'Gpx1/3ko', 7: 'Hg6pd', 8: 'BMPR2', 9: 'GPP29'}

#Initializes Win32com.client to access specific document and worksheet
wb = excel.Workbooks.Add('C:\Users\imthou\Desktop\\test.xls')
ws = wb.Sheets(mouse_breed)
excel.Visible = True

#references
used = ws.UsedRange
nrows = used.Row + used.Rows.Count + 1
ncols = used.Column + used.Columns.Count - 1


#writes data to excel file, two exceptions for BMPR2 and GPX1/3ko animals who have slightly different formatting
if mouse_breed == 8:
    for index, n in enumerate(BMPR2_males):
        ws.Cells(nrows + index, 1).Value = n
        ws.Cells(nrows + index, 4).Value = DOB
        ws.Cells(nrows + index, 5).Value = breed[mouse_breed]
        ws.Cells(nrows + index, 6).Value = damID
        ws.Cells(nrows + index, 7).Value = damGT
        ws.Cells(nrows + index, 8).Value = sireID
        ws.Cells(nrows + index, 9).Value = sireGT

elif mouse_breed == 6:
    for index, n in enumerate(Male_animals):
        ws.Cells(nrows + index, 1).Value = n
        ws.Cells(nrows + index, 5).Value = DOB
        ws.Cells(nrows + index, 6).Value = breed[mouse_breed]
        ws.Cells(nrows + index, 7).Value = damID
        ws.Cells(nrows + index, 8).Value = damGT
        ws.Cells(nrows + index, 9).Value = sireID
        ws.Cells(nrows + index, 10).Value = sireGT


    for index, n in enumerate(Female_animals):
        ws.Cells(nrows + index +len(Male_animals), 1).Value = n
        ws.Cells(nrows + index +len(Male_animals), 5).Value = DOB
        ws.Cells(nrows + index +len(Male_animals), 6).Value = breed[mouse_breed]
        ws.Cells(nrows + index +len(Male_animals), 7).Value = damID
        ws.Cells(nrows + index +len(Male_animals), 8).Value = damGT
        ws.Cells(nrows + index +len(Male_animals), 9).Value = sireID
        ws.Cells(nrows + index +len(Male_animals), 10).Value = sireGT

else:
    for index, n in enumerate(Male_animals):
        ws.Cells(nrows + index, 1).Value = n
        ws.Cells(nrows + index, 4).Value = DOB
        ws.Cells(nrows + index, 5).Value = breed[mouse_breed]
        ws.Cells(nrows + index, 6).Value = damID
        ws.Cells(nrows + index, 7).Value = damGT
        ws.Cells(nrows + index, 8).Value = sireID
        ws.Cells(nrows + index, 9).Value = sireGT
    for index, n in enumerate(Female_animals):
        ws.Cells(nrows + index +len(Male_animals), 1).Value = n
        ws.Cells(nrows + index +len(Male_animals), 4).Value = DOB
        ws.Cells(nrows + index +len(Male_animals), 5).Value = breed[mouse_breed]
        ws.Cells(nrows + index +len(Male_animals), 6).Value = damID
        ws.Cells(nrows + index +len(Male_animals), 7).Value = damGT
        ws.Cells(nrows + index +len(Male_animals), 8).Value = sireID
        ws.Cells(nrows + index +len(Male_animals), 9).Value = sireGT

