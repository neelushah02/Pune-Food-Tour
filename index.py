#opening the data sheet
import openpyxl
import math

#quick definition of distance between two points
def calculateDistance(x1,y1,x2,y2):
     dist = math.sqrt((x2 - x1)**2 + (y2 - y1)**2)
     return dist

#beginning processing
wb = openpyxl.load_workbook('C:/Users/neelu/OneDrive/Desktop/Computer Science/Pune Food Tour/place_data.xlsx')
main_sheet = wb.get_sheet_by_name('Sheet0') #original sheet exported from google maps

#for each location on the sheet we have to calculate distance to all other
#locations, these distances must be written to sheets in the workbook

for i in range(1 , 25): #len(wb['A'])+1
    first_x_location = 'B' + str(i)
    first_x = main_sheet[first_x_location].value
    first_y_location = 'C' + str(i)
    first_y = main_sheet[first_y_location].value
    #create sheet for the distances between all other destinations and i'th one
    name_of_new_sheet = 'Sheet' + str(i)
    wb.create_sheet(index = i, title = name_of_new_sheet)
    current_sheet = wb.get_sheet_by_name(name_of_new_sheet)

    for j in range(1 , 25): #len(wb['A'])+1
        second_x_location = 'B' + str(j)
        second_x = main_sheet[second_x_location].value
        second_y_location = 'C' + str(j)
        second_y = main_sheet[second_y_location].value
        temp = calculateDistance(float(first_x),float(first_y),float(second_x),float(second_y))
        current_cell = 'A' + str(j)
        from_cell = 'A' + str(i)
        #writing something like 'Home to Yewale'
        current_sheet.cell(row = j, column = 1).value = main_sheet[from_cell].value + ' to ' + main_sheet[current_cell].value
        #writing the distance
        current_cell = 'B' + str(j)
        current_sheet.cell(row = j, column = 2).value = temp

wb.save('C:/Users/neelu/OneDrive/Desktop/Computer Science/Pune Food Tour/place_data.xlsx')
