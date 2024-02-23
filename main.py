
from handle_data import GoogleSheetManager
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", 'https://www.googleapis.com/auth/spreadsheets']


def main():
    creds_file = "credentials.json"
    sheet_api = GoogleSheetManager(creds_file, SCOPES)
    sheet_api.build_service("sheets", "v4")
    spreadsheet = sheet_api.connect_spreadsheet('1u40t0Ui9Frn7N8kF7Q3r-PS2H2VT6lC0HoWQgwp1qYI')
    # # spreadsheet_id = spreadsheet.get('spreadsheetId')
    # # print(sheet_api.get_worksheet_by_title(spreadsheet.get('spreadsheetId'), 'Sheet1'))
    # sheet_api.create_worksheet(spreadsheet.get('spreadsheetId'), 'worksheet 2')
    # # sheet_api.create_spreadsheet("")
    # data = [
    #     ["STT",  'Họ', 'Tên', 'SDT', 'Ngày sinh'],
    #     [1, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    #     [2, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    #     [3, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    #     [4, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    #     [5, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    #     [6, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    # ]
    # # sheet_api.write_data_range(spreadsheet.get('spreadsheetId'), "'Sheet2'!A1:B2", data)
    data = [
        ['Team', 'Member', 'Number of tasks']
    ]

    format_cell = {
        "backgroundColor": {
            "red": 0.0,
            "green": 0.0,
            "blue": 0.0
        },
        "horizontalAlignment": "CENTER",
        "textFormat": {
            "foregroundColor": {
                "red": 1.0,
                "green": 1.0,
                "blue": 1.0
            },
            "fontSize": 12,
            "bold": True
        }
    }
    format_range = {
        "sheetId": None,
        "startRowIndex": 0,
        "endRowIndex": 1,
        "startColumnIndex": 0,
        "endColumnIndex": len(data[0])
    }

    merge_range = {
        "sheetId": '0',
        "startRowIndex": 0,
        "endRowIndex": 2,
        "startColumnIndex": 3,
        "endColumnIndex": 4
    }

    # sheet_api.import_data_to_google_sheet(data, format_cell, format_range,merge_range)

    # create BPO report

    # sheet_api.import_data_to_google_sheet(data, format_cell, format_range, merge_range)

    data = [
        {
            'team_name': 'Team #1',
            'members_name': ['NhuLTQ', 'UyenNHP', 'TramPTN', 'DungDNT'],
            'numbers_task': [
                {
                    'member_name': 'NhuLTQ',
                    'number_of_tasks': 81,
                    'resolve_hour_estimate': {
                        'time': 106.4,
                        'man_month': 3.01
                    },
                    'resolve_hour_actual': {
                        'time': 106.4,
                        'man_month': 3.01
                    },
                    'diff': {
                        'time': 106.4,
                        'man_month': 3.01
                    }
                },
                {
                    'member_name': 'UyenNHP',
                    'number_of_tasks': 81,
                    'resolve_hour_estimate': {
                        'time': 106.4,
                        'man_month': 3.01
                    },
                    'resolve_hour_actual': {
                        'time': 106.4,
                        'man_month': 3.01
                    },
                    'diff': {
                        'time': 106.4,
                        'man_month': 3.01
                    }
                }
            ]
        }
    ]
    spreadsheet_id = spreadsheet.get('spreadsheetId')
    # sheet_api.merge_cells(spreadsheet_id, merge_range, merge_type='MERGE_COLUMNS')
    # sheet_api.write_data_range(spreadsheet_id, range_name="'Sheet1'!A1", data=[[]])
    # merge_range = {
    #     "sheetId": '0',
    #     "startRowIndex": 0,
    #     "endRowIndex": 2,
    #     "startColumnIndex": 0,
    #     "endColumnIndex": 2
    # }
    # sheet_api.merge_cells(spreadsheet_id, merge_range='MERGE_COLUMNS')

    sheet_api.create_table_bpo(start_char=65, start_row_index=2, end_row_index=2, start_col_index=2, end_col_index=2)
    # sheet_api.write_data_range(spreadsheet_id, "'Sheet0'!A4:A3", data=[['a', 'c', 'd', 'e']])


main()

# start_char = 65
# sheet_id = 0

# start_row_index = 3
# end_row_index = 3
# start_col_index = 3
# end_col_index = 3

# init_cells = [
#     {
#         'col_name': [['Team']],
#         'merge_range': {
#             "sheetId": sheet_id,
#             "startRowIndex":  0 + start_row_index,
#             "endRowIndex":  2 + end_row_index,
#             "startColumnIndex":  0 + start_col_index,
#             "endColumnIndex":   1 + end_col_index
#         },
#         'merge_type': 'MERGE_COLUMNS',
#     },
#     {
#         'col_name': [['Member']],
#         'merge_range': {
#             "sheetId": sheet_id,
#             "startRowIndex":  0 + start_row_index,
#             "endRowIndex":  2 + end_row_index,
#             "startColumnIndex":  1 + start_col_index,
#             "endColumnIndex":   3 + end_col_index
#         },
#         'merge_type': 'MERGE_ALL',
#     },
#     {
#         'col_name': [["Number of tasks"]],
#         'merge_range': {
#             "sheetId": sheet_id,
#             "startRowIndex":  0 + start_row_index,
#             "endRowIndex":  2 + end_row_index,
#             "startColumnIndex":  3 + start_col_index,
#             "endColumnIndex":   4 + end_col_index
#         },
#         'merge_type': 'MERGE_COLUMNS',
#     },
#     {
#         'cells': [
#             {
#                 'col_name': [["Resolve hour estimate"]],
#                 'merge_range': {
#                     "sheetId": sheet_id,
#                     "startRowIndex":  0 + start_row_index,
#                     "endRowIndex":  1 + end_row_index,
#                     "startColumnIndex":  4 + start_col_index,
#                     "endColumnIndex":   6 + end_col_index
#                 },
#                 'merge_type': 'MERGE_ROWS',
#             },
#             {
#                 'col_name': [["Time", "Man+month"]],
#                 'merge_range': {
#                     "sheetId": sheet_id,
#                     "startRowIndex":  1 + start_row_index,
#                     "endRowIndex":  2 + end_row_index,
#                     "startColumnIndex":  4 + start_col_index,
#                     "endColumnIndex":   5 + end_col_index
#                 },
#                 'merge_type': 'MERGE_ROWS',
#             }
#         ]
#     },
#     {
#         'cells': [
#             {
#                 'col_name': [["Resolve hour actual"]],
#                 'merge_range': {
#                     "sheetId": sheet_id,
#                     "startRowIndex":  0 + start_row_index,
#                     "endRowIndex":  1 + end_row_index,
#                     "startColumnIndex":  6 + start_col_index,
#                     "endColumnIndex":   8 + end_col_index
#                 },
#                 'merge_type': 'MERGE_ROWS',
#             },
#             {
#                 'col_name': [["Time", "Man+month"]],
#                 'merge_range': {
#                     "sheetId": sheet_id,
#                     "startRowIndex":  1 + start_row_index,
#                     "endRowIndex":  2 + end_row_index,
#                     "startColumnIndex":  6 + start_col_index,
#                     "endColumnIndex":   7 + end_col_index
#                 },
#                 'merge_type': 'MERGE_ROWS',
#             }
#         ]
#     },
#     # {
#     #     'cells': [
#     #         {
#     #             'col_name': [["Resolve hour Diff"]],
#     #             'merge_range': {
#     #                 "sheetId": sheet_id,
#     #                 "startRowIndex": 0,
#     #                 "endRowIndex": 1,
#     #                 "startColumnIndex": 8,
#     #                 "endColumnIndex": 10
#     #             },
#     #             'merge_type': 'MERGE_ROWS',
#     #         },
#     #         {
#     #             'col_name': [["Time", "Man-month"]],
#     #             'merge_range': {
#     #                 "sheetId": sheet_id,
#     #                 "startRowIndex": 1,
#     #                 "endRowIndex": 2,
#     #                 "startColumnIndex": 8,
#     #                 "endColumnIndex": 9
#     #             },
#     #             'merge_type': 'MERGE_ROWS',
#     #         }
#     #     ]
#     # },

# ]

# # for value in init_cells:
# #     if value.get('cells', None) is not None:
# #         for cell in value['cells']:
# #             range_name = f"'Sheet1'!{chr(start_char + cell['merge_range']['startColumnIndex'])}{(cell['merge_range']['startRowIndex'] + 1)}"
# #             sheet_api.merge_cells(spreadsheet_id, merge_range=cell['merge_range'], merge_type=cell['merge_type'])
# #             sheet_api.write_data_range(spreadsheet_id, range_name=range_name, data=cell['col_name'])
# #     else:
# #         sheet_api.merge_cells(spreadsheet_id, merge_range=value['merge_range'], merge_type=value['merge_type'])
# #         range_name = f"'Sheet1'!{chr(start_char + value['merge_range']['startColumnIndex'])}{(value['merge_range']['startRowIndex'] + 1)}"
# #         sheet_api.write_data_range(spreadsheet_id, range_name=range_name, data=value['col_name'])
