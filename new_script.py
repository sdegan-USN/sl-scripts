from docx import Document
from openpyxl import Workbook

# Load the Word document
document = Document('test.docx')

# Create a new Excel workbook
workbook = Workbook()
worksheet = workbook.active


# Table 1
worksheet.append(['Table 1'])
worksheet.append(['Macrostate', 'Possible Microstates(Dice Combinations)', 'Number of Microstates, Ω', 'Entropy S= k ln(Ω)'])
worksheet.append(['2', '\xa01 black and 1 white', '1\xa0', '4.8 x 10^-24\xa0'])
worksheet.append(['3', '\xa0', '2', '9.6 x 10-24'])
worksheet.append(['4', '\xa0(1,3), (2,2), (3,1)', '\xa03', '15.4 x 10^-24\xa0'])
worksheet.append(['5', '\xa0(1,4), (2,3), (3,2), (4,1)', '\xa04', '\xa019.2 x 10^-24'])
worksheet.append(['6', '\xa0(1, 5), (2,4), (3,3), (4,2), (5, 1)', '\xa05', '\xa025 x 10^-24'])
worksheet.append(['7', '\xa0(1, 6), (2, 5), (3,4), (4,3), (5, 2), (6,1)', '\xa06', '\xa030.8 x 10^-24'])
worksheet.append(['8', '\xa0(2,6), (3,5), (4,4), (5,3), (6,2)', '\xa05', '\xa025 x 10^-24'])
worksheet.append(['9', '\xa0(3, 6), (4,5), (5,4), (6,3)', '\xa04', '\xa019.2 x 10^-24'])
worksheet.append(['10', '\xa0(4,6), (5,5), (6,4)', '\xa03', '\xa015.4 x 10^-24'])
worksheet.append(['11', '\xa0(5,6), (6,5)', '\xa02', '\xa09.6 x 10^-24'])
worksheet.append(['12', '\xa0(6,6)', '\xa01', '\xa04.8 x 10^-24'])
worksheet.append([])

# Table 2
worksheet.append(['Table 2'])
worksheet.append(['Macrostate', 'Probability of Rolling a Macrostate'])
worksheet.append(['2', '\xa02.8%'])
worksheet.append(['3', '\xa05.6%'])
worksheet.append(['4', '\xa08.3%'])
worksheet.append(['5', '\xa011.1%'])
worksheet.append(['6', '\xa013.9%'])
worksheet.append(['7', '\xa016.7%'])
worksheet.append(['8', '\xa013.9%'])
worksheet.append(['9', '\xa011.1%'])
worksheet.append(['10', '\xa08.3%'])
worksheet.append(['11', '\xa05.6%'])
worksheet.append(['12', '\xa02.8%'])
worksheet.append([])

# Table 3
worksheet.append(['Table 3'])
worksheet.append(['Initial Temperature (ºC)', '24.4'])
worksheet.append(['Final Temperature (ºC)', '24.4 is initial, 29.8 was final (upper box not working)'])
worksheet.append([])

# Table 4
worksheet.append(['Table 4'])
worksheet.append(['Macrostate', 'Number of Occurrences', 'Total Occurrences'])
worksheet.append(['2', 'lll', '\xa03'])
worksheet.append(['3', '\xa0lllll', '\xa05'])
worksheet.append(['4', '\xa0llll', '\xa04'])
worksheet.append(['5', '\xa0lllll ll ', '\xa07'])
worksheet.append(['6', '\xa0lllll llll', '\xa09'])
worksheet.append(['7', '\xa0lllll lllll lllll l', '\xa016'])
worksheet.append(['8', '\xa0lllll lllll llll ', '\xa014'])
worksheet.append(['9', '\xa0lllll lllll lllll ll', '\xa017'])
worksheet.append(['10', '\xa0lllll lllll ll', '\xa012'])
worksheet.append(['11', '\xa0lllll lllll', '\xa09'])
worksheet.append(['12', '\xa0llll', '\xa04'])
worksheet.append([])
# Save the Excel workbook
workbook.save('tables.xlsx')