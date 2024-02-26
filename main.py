
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

    # sheet_api.create_table_for_bpo()
    # sheet_api.write_data_range(spreadsheet_id, "'Sheet0'!A4:A3", data=[['a', 'c', 'd', 'e']])

    # values = sheet_api.get_values_range('1xzUdAkGT8YBbM-v0aVqkNad7262IlKo9Uuuv-Ddhz50', "'Sheet0'!A1:J17")
    # print(values)

    # data = {
    #         'range': 'Sheet0!A1:J17',
    #         'majorDimension': 'ROWS',
    #         'values': [
    #             ['Team', 'Member', '', 'Number of tasks', 'Resolve hour estimate', '', 'Resolve hour actual', '', 'Diff'],
    #             ['', '', '', '', 'Time', 'Man-month', 'Time', 'Man-month', 'Time', 'Man-month'],
    #             ['Team #1', 'Total', '', '237', '361,20', '3,01', '400,40', '3,34', '-39,2', '-0,33'],
    #             ['', 'Team', 'NhuLTQ', '81', '106,4', '0,89', '107,85', '0,90', '-1,45', '-0,01'],
    #             ['', '', 'UyenNHP', '64', '109,8', '0,92', '140,8', '1,17', '-31', '-0,26'],
    #             ['', '', 'TramPTN', '43', '65', '0,54', '73,25', '0,61', '-8,25', '-0,07'],
    #             ['', '', 'DungDNT', '49', '80', '0,67', '78,5', '0,65', '1,5', '0,01'],
    #             ['Team #2', 'Total', '', '242', '468,9', '3,91', '485,9', '4,05', '-17', '-0,14'],
    #             ['', 'Team', 'LinhPTT', '68', '123,15', '1,03', '129,45', '1,08', '-6,3', '-0,05'],
    #             ['', '', 'DongCTM', '62', '116,75', '0,97', '107,75', '0,90', '9', '0,08'],
    #             ['', '', 'TuyenHT', '70', '130', '1,08', '138', '1,15', '-8', '-0,07'],
    #             ['', '', 'ChiNTK', '42', '99', '0,83', '110,7', '0,92', '-11,7', '-0,10'],
    #             ['Team #3', 'Total', '', '127', '272,75', '2,27', '249,5', '2,08', '23,25', '0,19'],
    #             ['', 'Team', 'TienLT', '77', '130,5', '1,09', '112', '0,93', '18,5', '0,15'],
    #             ['', '', 'PhuongNTT', '35', '78,25', '0,65', '69,5', '0,58', '8,75', '0,07'],
    #             ['', '', 'LuanND', '15', '64', '0,53', '68', '0,57', '-4', '-0,03'],
    #             ['Total', '', '', '606', '1.102,85', '9,19', '1.135,80', '9,47', '-32,95', '-0,27']
    #         ]
    #     }
    data_json = [
        {
            "team_name": "Team #1",
            'quantity_member': 5,
            'task_members': [
                {
                    'name_member': 'NhuLTQ',
                    'number_of_tasks': 0,
                    'resolve_hour_estimate': {
                        'time': 1,
                        'man_month': 2
                    },
                    'resolve_hour_actual': {
                        'time': 3,
                        'man_month': 4
                    },
                    'diff': {
                        'time': 5,
                        'man_month': 6
                    },
                },
                {
                    'name_member': 'NhuLTQ 2',
                    'number_of_tasks': 7,
                    'resolve_hour_estimate': {
                        'time': 8,
                        'man_month': 9
                    },
                    'resolve_hour_actual': {
                        'time': 10,
                        'man_month': 11
                    },
                    'diff': {
                        'time': 12,
                        'man_month': 13
                    },
                },
                {
                    'name_member': 'NhuLTQ 3',
                    'number_of_tasks': 14,
                    'resolve_hour_estimate': {
                        'time': 15,
                        'man_month': 16
                    },
                    'resolve_hour_actual': {
                        'time': 17,
                        'man_month': 18
                    },
                    'diff': {
                        'time': 19,
                        'man_month': 20
                    },
                },
                {
                    'name_member': 'NhuLTQ 4',
                    'number_of_tasks': 21,
                    'resolve_hour_estimate': {
                        'time': 22,
                        'man_month': 23
                    },
                    'resolve_hour_actual': {
                        'time': 24,
                        'man_month': 25
                    },
                    'diff': {
                        'time': 26,
                        'man_month': 27
                    },
                },
                {
                    'name_member': 'NhuLTQ 4',
                    'number_of_tasks': 28,
                    'resolve_hour_estimate': {
                        'time': 29,
                        'man_month': 30
                    },
                    'resolve_hour_actual': {
                        'time': 31,
                        'man_month': 32
                    },
                    'diff': {
                        'time': 33,
                        'man_month': 34
                    },
                },
            ]
        },
        {
            "team_name": "Team #1",
            'quantity_member': 5,
            'task_members': [
                {
                    'name_member': 'NhuLTQ',
                    'number_of_tasks': 35,
                    'resolve_hour_estimate': {
                        'time': 36,
                        'man_month': 37
                    },
                    'resolve_hour_actual': {
                        'time': 38,
                        'man_month': 39
                    },
                    'diff': {
                        'time': 40,
                        'man_month': 41
                    },
                },
                {
                    'name_member': 'NhuLTQ 2',
                    'number_of_tasks': 42,
                    'resolve_hour_estimate': {
                        'time': 43,
                        'man_month': 44
                    },
                    'resolve_hour_actual': {
                        'time': 45,
                        'man_month': 46
                    },
                    'diff': {
                        'time': 47,
                        'man_month': 48
                    },
                },
                {
                    'name_member': 'NhuLTQ 3',
                    'number_of_tasks': 49,
                    'resolve_hour_estimate': {
                        'time': 50,
                        'man_month': 51
                    },
                    'resolve_hour_actual': {
                        'time': 52,
                        'man_month': 53
                    },
                    'diff': {
                        'time': 54,
                        'man_month': 55
                    },
                },
                {
                    'name_member': 'NhuLTQ 4',
                    'number_of_tasks': 56,
                    'resolve_hour_estimate': {
                        'time': 57,
                        'man_month': 58
                    },
                    'resolve_hour_actual': {
                        'time': 59,
                        'man_month': 60
                    },
                    'diff': {
                        'time': 61,
                        'man_month': 62
                    },
                },
                {
                    'name_member': 'NhuLTQ 4',
                    'number_of_tasks': 63,
                    'resolve_hour_estimate': {
                        'time': 64,
                        'man_month': 65
                    },
                    'resolve_hour_actual': {
                        'time': 65,
                        'man_month': 67
                    },
                    'diff': {
                        'time': 68,
                        'man_month': 69
                    },
                },
            ]
        },
        {
            "team_name": "Team #2",
            'quantity_member': 3,
            'task_members': [
                {
                    'name_member': 'TaiLTQ',
                    'number_of_tasks': 70,
                    'resolve_hour_estimate': {
                        'time': 71,
                        'man_month': 72
                    },
                    'resolve_hour_actual': {
                        'time': 73,
                        'man_month': 74
                    },
                    'diff': {
                        'time': 75,
                        'man_month': 76
                    },
                },
                {
                    'name_member': 'TaiLTQ 2',
                    'number_of_tasks': 77,
                    'resolve_hour_estimate': {
                        'time': 78,
                        'man_month': 79
                    },
                    'resolve_hour_actual': {
                        'time': 80,
                        'man_month': 81
                    },
                    'diff': {
                        'time': 82,
                        'man_month': 83
                    },
                },
                {
                    'name_member': 'TaiLTQ 3',
                    'number_of_tasks': 84,
                    'resolve_hour_estimate': {
                        'time': 85,
                        'man_month': 86
                    },
                    'resolve_hour_actual': {
                        'time': 87,
                        'man_month': 88
                    },
                    'diff': {
                        'time': 89,
                        'man_month': 90
                    },
                },

            ]
        },

    ]

    sheet_api.create_table_for_bpo(data_json)


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
