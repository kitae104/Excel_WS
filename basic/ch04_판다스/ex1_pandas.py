import openpyxl
import pandas 

wb = openpyxl.load_workbook("writeonlyfile.xlsx")
sheet = wb.active

df = pandas.read_excel("writeonlyfile.xlsx", sheet_name='Sheet')
Age = df['Age']
print(Age)

age_sum = Age.sum()
print(age_sum)

age_mean = Age.mean()
print(age_mean)

sheet['D1'].value = age_sum
sheet['D2'].value = age_mean

wb.save("pandas.xlsx")