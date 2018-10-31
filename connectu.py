import xlsxwriter
import os.path
import random
assert os.path.isfile('responses1.xlsx')

# Creating class to import response information from excel sheet
from xlrd import open_workbook

class Arm(object):
    def __init__(self, id, time, first_name, last_name, email , phone_num, year,gender, house):
        self.id = id
        self.timestamp = time
        self.First_Name = first_name
        self.Name = last_name
        self.Email = email
        self.Phone_Number = phone_num
        self.Class_Year = year
        self.Gender = gender
        self.House = house

    def __str__(self):
        return("Arm object:\n"
               "  ID = {0}\n"
               "  First Name = {1}\n"
               "  Last Name = {2}\n"
               "  Email = {3}\n"
               "  Phone Number = {4} \n"
               "  Class Year = {5} \n"
               "  Gender = {6} \n"
               "  House = {7} \n"
               .format(self.id, self.First_Name, self.Name,
                       self.Email, self.Phone_Number, self.Class_Year,
                       self.Gender, self.House));


# importing in ConnectU information from excel sheet
wb = open_workbook('connecturesponsesnov192017.xlsx')  # place file name that you want read Here
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols
    items = []
    rows = []
    counter = 0
    for row in range(1, number_of_rows):
        values = []
        counter = counter + 1
        value = counter
        values.append(value)
        for col in range(number_of_columns-1):
            value  = (sheet.cell(row,col).value)
            try:
                value = str(value)
            except ValueError:
                pass
            finally:
                values.append(value)
        item = Arm(*values)
        items.append(item)
# marks the previous matches from last week to not allow people to not get same person
prev_matchlist = [(9,4),(15,17),(1,31),(24,34),(29,26),(32,30),(3,16),(18,10),(19,12),(21,5),(8,2),(28,22),(11,7),(35,6),(20,25),(27,23),(33,13),(14,17)]

# randomizing the list of items
random.shuffle(items)

# printing out the information
#for item in items:
 #   print item
  #  print("Accessing one single value (eg. DSPName): {0}".format(item.First_Name))
  #  print

# function that checks to see if person has already been matched up with someone else
def alreadymatched(matchlst,(pair1,pair2)):
    for i in range(len(matchlst)):
        if pair1 in matchlst[i] or pair2 in matchlst[i] : # seeing if either the first or second person has already been paired
            return True
        else:
            if i == len(matchlst)-1: # if function reaches the end of the list without seeing that anyone has been matched it
                return False

def previouslymatched(prev_matchlst,(pair1,pair2)):
    if ((pair1,pair2) in prev_matchlst) or ((pair2,pair1) in prev_matchlst):
        return True
    else:
        return False

#matching people up
id_matches = []
matches = []
itemlngth = len(items)

for item1 in items:
    counter = 0
    indicator = 0
    item3 = None
    item4 = None
    for item2 in items:
        counter = counter + 1
        if (not alreadymatched(id_matches,(item1.id,item2.id))) and item1.id != item2.id \
                and (not previouslymatched(prev_matchlist,(item1.id,item2.id))):
            # ensuring mixture of class years
            if int(float(item1.Class_Year)) != int(float(item2.Class_Year)):
                id_matches.append((item1.id,item2.id))
                matches.append((item1,item2))
                break
            else:
                if indicator == 0:
                    indicator = 1
                    item3 = item1
                    item4 = item2
        if itemlngth == counter and indicator == 1:
            id_matches.append((item3.id,item4.id))
            matches.append((item3 , item4))
        if itemlngth == counter and indicator == 0 and not alreadymatched(id_matches,(item1.id,item2.id)):
            print("person not matched {0}\n"
                  "ID = {1}"
                  .format(item1.First_Name,item1.id))



# creating the excel sheet with matched up people

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('demomatchup3.xlsx')
worksheet = workbook.add_worksheet()


# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Writing headers of file
worksheet.write('A1', 'First Name');
worksheet.write('B1', 'Last Name')
worksheet.write('C1', 'Email')
worksheet.write('D1', 'Phone Number')
worksheet.write('E1', 'Class Year')
worksheet.write('F1', 'Gender')
worksheet.write('G1', 'House')
worksheet.write('H1', 'Partner First Name')
worksheet.write('I1', 'Partner Last Name')
worksheet.write('J1', 'Partner Email')
worksheet.write('K1', 'Partner Phone Number')
worksheet.write('L1', 'Partner Class Year')
worksheet.write('M1', 'Partner Gender')
worksheet.write('N1', 'Partner House')
worksheet.write('O1', 'ID Number')

# Write some numbers, with row/column notation.
counter = -1
for (item1,item2) in matches:
    counter = counter +2
    worksheet.write(counter, 0, item1.First_Name)
    worksheet.write(counter+1, 0, item2.First_Name)
    worksheet.write(counter, 1, item1.Name)
    worksheet.write(counter + 1, 1, item2.Name)
    worksheet.write(counter, 2, item1.Email)
    worksheet.write(counter + 1, 2, item2.Email)
    worksheet.write(counter, 3, int(float(item1.Phone_Number)))
    worksheet.write(counter + 1, 3, int(float(item2.Phone_Number)))
    worksheet.write(counter, 4, int(float(item1.Class_Year)))
    worksheet.write(counter + 1, 4, int(float(item2.Class_Year)))
    worksheet.write(counter, 5, item1.Gender)
    worksheet.write(counter + 1, 5, item2.Gender)
    worksheet.write(counter, 6, item1.House)
    worksheet.write(counter + 1, 6, item2.House)

    worksheet.write(counter, 7, item2.First_Name)
    worksheet.write(counter + 1, 7, item1.First_Name)
    worksheet.write(counter, 8, item2.Name)
    worksheet.write(counter + 1, 8, item1.Name)
    worksheet.write(counter, 9, item2.Email)
    worksheet.write(counter + 1, 9, item1.Email)
    worksheet.write(counter, 10, int(float(item2.Phone_Number)))
    worksheet.write(counter + 1, 10, int(float(item1.Phone_Number)))
    worksheet.write(counter, 11, int(float(item2.Class_Year)))
    worksheet.write(counter + 1, 11, int(float(item1.Class_Year)))
    worksheet.write(counter, 12, item2.Gender)
    worksheet.write(counter + 1, 12, item1.Gender)
    worksheet.write(counter, 13, item2.House)
    worksheet.write(counter + 1, 13, item1.House)
    worksheet.write(counter, 14, item1.id)
    worksheet.write(counter+1, 14, item2.id)
workbook.close()

