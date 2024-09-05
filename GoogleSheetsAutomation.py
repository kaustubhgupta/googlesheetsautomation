import gspread
import json

with open("creds.json") as f:
    credentials = json.load(f)

gc = gspread.service_account_from_dict(credentials)

## Creating and sharing the new spreadsheet
# sh = gc.create('ArticleDemoTest')
# sh.share(email_address='kaustubhgupta1828@gmail.com', perm_type='user', role='writer', notify=True, email_message="This is a test file")


## Different ways to open the spreadsheet
# sh = gc.open("ArticleDemo")
# sh = gc.open_by_key("1R97twcM0FfFNSsrh_0FjDDg-HcQF5PLHbhRxu9pTV_Q")
# sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1R97twcM0FfFNSsrh_0FjDDg-HcQF5PLHbhRxu9pTV_Q/edit?gid=0#gid=0")
# print(sh.sheet1.acell('A1').value)

## Different ways to select a worksheet
# print(sh.get_worksheet(0))
# print(sh.worksheet("ArticleWorkSheet1"))
# print(sh.sheet1)
# print(sh.get_worksheet_by_id(0))

# print("Now fetching all sheets...")

## Returning all worksheets
# for ws in sh.worksheets():
#     print(ws)

## Adding new worksheets to the existing spreadsheet
# sh.add_worksheet('ArticleWorkSheet1.5', rows=100, cols=20, index=1)

## Renaming worksheet titles
# print(sh.worksheet("ArticleWorkSheet3").update_title("ArticleWorkSheet2.5"))

## Deleting worksheets
# sh.del_worksheet(sh.worksheet("ArticleWorkSheet2.5"))
# sh.del_worksheet_by_id('602396579')

## Cell properties
# sampleCell = sh.worksheet("ArticleWorkSheet1").cell(row=1, col=1)
# print('Row: {}\nColumn: {}\nValue: {}\nAddress: {}'.format(sampleCell.row, sampleCell.col, sampleCell.value, sampleCell.address))

## Insert row operations
# sampleWorksheet = sh.sheet1
# sampleWorksheet.insert_row(
#     ['A', 'B', 'C', 'D']
# )
# sampleWorksheet.insert_rows(
#     [
#         ['KG', 54, 23, 12],
#         ['OG', 34, 12, 34],
#         ['ME', 23, 45, 90],
#         ['YE', 65, 12, 54]
#     ], row=2
# )

## Appending rows and column insertion
# sampleWorksheet.append_rows(
#     [
#         ['SN', 67, 87, 45],
#         ['AR', 56, 23, 65]
#     ],
#     table_range="A1:D5"
# )
# sampleWorksheet.insert_cols(
#     [
#         ['E', 56, 34, 65, 34, 76, 45]
#     ],
#     col=5
# )


## Fetching single cell
# print(sampleWorksheet.acell('A1').row)
# print(sampleWorksheet.cell(1, 1).value)

## Fetching all Cells of the Worksheet or Range
# print(sampleWorksheet.get_all_cells())
# print(sampleWorksheet.range('B4:E5'))

## Fetching all Values from a Row or Column
# print(sampleWorksheet.row_values(1))
# print(sampleWorksheet.col_values(4))

## Fetching range of cells
# print('Get Range: {}'.format(sampleWorksheet.get("A1:D4")))
# print('Batch Get Range: {}'.format(sampleWorksheet.batch_get([
#     "A1:D4",
#     "B4:E3"
# ])))

# import pandas as pd
# print(pd.DataFrame(sampleWorksheet.get_all_records()))
# print(pd.DataFrame(sampleWorksheet.get_all_values()))

# print(sampleWorksheet.get_all_cells())

## Updating single cells
# print(sampleWorksheet.update_acell('A2', 'Kaustubh'))
# print(sampleWorksheet.update_acell('A3', 'Oggy'))
# print(sampleWorksheet.update([['Hello']], 'A4'))

## Updating range of cells
# rangeOfCells = sampleWorksheet.range('B2:B7')
# for cell in rangeOfCells:
#     newValue = int(cell.value) + 10
#     cell.value = newValue
# print(sampleWorksheet.update_cells(rangeOfCells))

## Updating multiple range of cells
# range1 = 'C2:C7'
# range2 = 'E2:E7'
# bothRangeValues = sampleWorksheet.batch_get([
#     range1,
#     range2
# ])
# range1Values, range2Values = bothRangeValues
# range1UpdatedValues = [[int(x[0]) + 10] for x in range1Values]
# range2UpdatedValues = [[int(x[0]) + 20] for x in range2Values]
# print(sampleWorksheet.batch_update([
#     {
#         'range': range1,
#         'values': range1UpdatedValues
#     },

#     {
#         'range': range2,
#         'values': range2UpdatedValues
#     }
# ]))

## Deleting rows and columns
# print(sampleWorksheet.delete_columns(4))
# print(sampleWorksheet.delete_rows(6))

## Searching cells
# import re
# print(sampleWorksheet.find('64', in_column=2))
# searchRe = re.compile(r'(a|A)')
# print(sampleWorksheet.findall(searchRe))

## Single formatting
# borderFormatting = {
#     "style": "SOLID",
#     "colorStyle": {"rgbColor": {"red": 0, "green": 0, "blue": 0, "alpha": 1}},
# }

# print(
#     sampleWorksheet.format(
#         "A1:D6",
#         format={
#             "borders": {
#                 "top": borderFormatting,
#                 "bottom": borderFormatting,
#                 "left": borderFormatting,
#                 "right": borderFormatting,
#             },
#         },
#     )
# )

# Batch formatting
# borderFormatting = {
#     "style": "SOLID",
#     "colorStyle": {"rgbColor": {"red": 0, "green": 0, "blue": 0, "alpha": 1}},
# }
# formats = [
#     {
#         "range": "A1:D6",
#         "format": {
#             "borders": {
#                 "top": borderFormatting,
#                 "bottom": borderFormatting,
#                 "left": borderFormatting,
#                 "right": borderFormatting,
#             },
#             "horizontalAlignment": "CENTER",
#         },
#     },
#     {
#         "range": "A1:D1",
#         "format": {
#             "textFormat": {
#                 "bold": True,
#             },
#             "backgroundColorStyle": {
#                 "rgbColor": {"red": 0.8, "green": 0.8, "blue": 1, "alpha": 0.8}
#             },
#         },
#     },
# ]

# print(sampleWorksheet.batch_format(formats))

## Clear a range from worksheet
# print(sampleWorksheet.batch_clear(["C1:C6"]))


## Clear entire worksheet
# print(sampleWorksheet.clear())
