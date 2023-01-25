import openpyxl
import pprint
wb = openpyxl.load_workbook('/workspaces/vaktaplan/dagskipan_Beta.xlsx')
ws = wb.active

yoff = [(i * 43) for i in range(7)]
weekday = ('B',1)
date = ('D',4)
mv_rows = [(8,9), (11,12), (14,15)]
dv_rows = [(17,18), (20,21), (23,24)]
nv_rows = [26,27]

cols = {
    'LRL': 'B',
    'PHA': 'D',
    'AS': 'F',
    'GR': 'H',
    'BEG': 'J',
    'GH': 'L'
}
pprint.pprint([ws[f'{weekday[0]}{weekday[1] + y}'].value for y in yoff])