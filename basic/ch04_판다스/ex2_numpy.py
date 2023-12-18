import openpyxl
import numpy 

wb = openpyxl.load_workbook("writeonlyfile.xlsx")
sheet = wb.active

age_values = [cell.value for cell in sheet['B']]
print(age_values[1:])
array_of_ages = numpy.array(age_values[1:])

sum = numpy.sum(array_of_ages)
print(sum)
mean = numpy.mean(array_of_ages)
print(mean)

sheet['E1'].value = sum
sheet['E2'].value = mean

wb.save("numpy.xlsx")