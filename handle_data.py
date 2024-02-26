
from google_sheet_api import GoogleSheetApi


class GoogleSheetManager(GoogleSheetApi):
    def __init__(self, creds_file, scopes) -> None:
        super().__init__(creds_file, scopes)

    def import_data_to_google_sheet(self, data: list[list], format_cell: dict, format_range: dict, merge_range: dict) -> None:
        """
            Import data to google drive
        Args:
            data (list[list]): Data to be imported
            cell_format (dict): format cell
            format_cell (dict): dict format
        """

        # Create spreadsheet
        spreadsheet = self.create_spreadsheet("New Spreadsheet")
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        # Create worksheet
        worksheet_title = 'New Worksheet'
        sheet_id = self.create_worksheet(spreadsheet_id, worksheet_title)

        # Format
        format_range['sheetId'] = sheet_id
        merge_range['sheetId'] = sheet_id
        self.format_header(spreadsheet_id, format_cell, format_range)

        # Tạo range dựa trên số dòng và cột
        num_rows = len(data)
        num_columns = len(data[0])
        range_name = f"'{worksheet_title}'!{chr(65)}1:{chr(65 + num_columns - 1)}{num_rows}"
        # Write data
        self.merge_cell_and_write_data(spreadsheet_id, merge_range, range_name, data)

    def create_table_bpo(self, start_char=65, start_row_index=0, end_row_index=0, start_col_index=0, end_col_index=0):
        """
            Create a table report BPO
        Args:
            start_char (int, optional): Kí tự bắt đầu table. Defaults to 65.
            start_row_index (int, optional): Chỉ số bắt đầu row. Defaults to 0.
            end_row_index (int, optional): Chỉ số end row. Defaults to 0.
            start_col_index (int, optional): Chỉ số bắt đầu col. Defaults to 0.
            end_col_index (int, optional): Chỉ số end row. Defaults to 0.
        """

        # Create spreadsheet
        spreadsheet_title = "BPO Report"
        spreadsheet = self.create_spreadsheet(spreadsheet_title)
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        # Create worksheet
        worksheet_title = 'Sheet1'
        sheet_id = self.create_worksheet(spreadsheet_id, worksheet_title)

        # Cell header
        cells_header = [
            {
                'col_name': [['Team']],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  0 + start_row_index,
                    "endRowIndex":  2 + end_row_index,
                    "startColumnIndex":  0 + start_col_index,
                    "endColumnIndex":   1 + end_col_index
                },
                'merge_type': 'MERGE_COLUMNS',
            },
            {
                'col_name': [['Member']],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  0 + start_row_index,
                    "endRowIndex":  2 + end_row_index,
                    "startColumnIndex":  1 + start_col_index,
                    "endColumnIndex":   3 + end_col_index
                },
                'merge_type': 'MERGE_ALL',
            },
            {
                'col_name': [["Number of tasks"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  0 + start_row_index,
                    "endRowIndex":  2 + end_row_index,
                    "startColumnIndex":  3 + start_col_index,
                    "endColumnIndex":   4 + end_col_index
                },
                'merge_type': 'MERGE_COLUMNS',
            },
            {

                'col_name': [["Resolve hour estimate"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  0 + start_row_index,
                    "endRowIndex":  1 + end_row_index,
                    "startColumnIndex":  4 + start_col_index,
                    "endColumnIndex":   6 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            },
            {
                'col_name': [["Time", "Man+month"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  1 + start_row_index,
                    "endRowIndex":  2 + end_row_index,
                    "startColumnIndex":  4 + start_col_index,
                    "endColumnIndex":   5 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            },
            {
                'col_name': [["Resolve hour actual"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  0 + start_row_index,
                    "endRowIndex":  1 + end_row_index,
                    "startColumnIndex":  6 + start_col_index,
                    "endColumnIndex":   8 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            },
            {
                'col_name': [["Time", "Man+month"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex":  1 + start_row_index,
                    "endRowIndex":  2 + end_row_index,
                    "startColumnIndex":  6 + start_col_index,
                    "endColumnIndex":   7 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            },
            {
                'col_name': [["Resolve hour Diff"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex": 0 + start_row_index,
                    "endRowIndex": 1 + end_row_index,
                    "startColumnIndex": 8 + start_col_index,
                    "endColumnIndex": 10 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            },
            {
                'col_name': [["Time", "Man-month"]],
                'merge_range': {
                    "sheetId": sheet_id,
                    "startRowIndex": 1 + start_row_index,
                    "endRowIndex": 2 + end_row_index,
                    "startColumnIndex": 8 + start_col_index,
                    "endColumnIndex": 9 + end_col_index
                },
                'merge_type': 'MERGE_ROWS',
            }
        ]

        # cells_body = [
        #     [{
        #         'col_name': [['Team #1']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  0 + start_col_index,
        #             "endColumnIndex":   1 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ALL',
        #     },
        #         {
        #         'col_name': [['Total']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  3 + end_row_index,
        #             "startColumnIndex":  1 + start_col_index,
        #             "endColumnIndex":   3 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ROWS',
        #     },
        #         {
        #         'col_name': [['Team']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  1 + start_col_index,
        #             "endColumnIndex":   2 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ALL',
        #     },
        #         {
        #         'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  2 + start_col_index,
        #             "endColumnIndex":   3 + end_col_index
        #         },
        #         'merge_type': None,
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  3 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ROWS',
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  4 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  4 + start_row_index,
        #             "endRowIndex":  5 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  5 + start_row_index,
        #             "endRowIndex":  6 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  6 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },], [{
        #         'col_name': [['Team #1']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  0 + start_col_index,
        #             "endColumnIndex":   1 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ALL',
        #     },
        #         {
        #         'col_name': [['Total']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  3 + end_row_index,
        #             "startColumnIndex":  1 + start_col_index,
        #             "endColumnIndex":   3 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ROWS',
        #     },
        #         {
        #         'col_name': [['Team']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  1 + start_col_index,
        #             "endColumnIndex":   2 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ALL',
        #     },
        #         {
        #         'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  2 + start_col_index,
        #             "endColumnIndex":   3 + end_col_index
        #         },
        #         'merge_type': None,
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  2 + start_row_index,
        #             "endRowIndex":  3 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': 'MERGE_ROWS',
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  3 + start_row_index,
        #             "endRowIndex":  4 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  4 + start_row_index,
        #             "endRowIndex":  5 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  5 + start_row_index,
        #             "endRowIndex":  6 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },
        #         {
        #         'col_name': [[237, 237, 237, 237, 237, 237, 237]],
        #         'merge_range': {
        #             "sheetId": sheet_id,
        #             "startRowIndex":  6 + start_row_index,
        #             "endRowIndex":  7 + end_row_index,
        #             "startColumnIndex":  3 + start_col_index,
        #             "endColumnIndex":   4 + end_col_index
        #         },
        #         'merge_type': "MERGE_ROWS",
        #     },]
        #     # C4:C7
        # ]
        # # Create cell header

        # cells_body = [
        #     {
        #         'teams': [
        #             {
        #                 'col_name': [['Team #1']],
        #                 'merge_range': {
        #                     "sheetId": sheet_id,
        #                     "startRowIndex":  2 + start_row_index,
        #                     "endRowIndex":  7 + end_row_index,
        #                     "startColumnIndex":  0 + start_col_index,
        #                     "endColumnIndex":   1 + end_col_index
        #                 },
        #                 'merge_type': 'MERGE_ALL',
        #             },
        #             {
        #                 'col_name': [['Team #2']],
        #                 'merge_range': {
        #                     "sheetId": sheet_id,
        #                     "startRowIndex":  2 + start_row_index,
        #                     "endRowIndex":  7 + end_row_index,
        #                     "startColumnIndex":  0 + start_col_index,
        #                     "endColumnIndex":   1 + end_col_index
        #                 },
        #                 'merge_type': 'MERGE_ALL',
        #             },
        #             {
        #                 'col_name': [['Team #3']],
        #                 'merge_range': {
        #                     "sheetId": sheet_id,
        #                     "startRowIndex":  2 + start_row_index,
        #                     "endRowIndex":  7 + end_row_index,
        #                     "startColumnIndex":  0 + start_col_index,
        #                     "endColumnIndex":   1 + end_col_index
        #                 },
        #                 'merge_type': 'MERGE_ALL',
        #             },
        #         ]
        #     },
        #     {
        #         'members': [
        #             {
        #                 'totals': [
        #                     {
        #                         'col_name': [['Total']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  2 + start_row_index,
        #                             "endRowIndex":  3 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                     {
        #                         'col_name': [['Total']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  2 + start_row_index,
        #                             "endRowIndex":  3 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                     {
        #                         'col_name': [['Total']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  2 + start_row_index,
        #                             "endRowIndex":  3 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                 ]
        #             },
        #             {
        #                 'teams': [
        #                     {
        #                         'col_name': [['Team']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   2 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                     {
        #                         'col_name': [['Team']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   2 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                     {
        #                         'col_name': [['Team']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  1 + start_col_index,
        #                             "endColumnIndex":   2 + end_col_index
        #                         },
        #                         'merge_type': 'MERGE_ALL',
        #                     },
        #                 ]
        #             },
        #             {
        #                 'names': [
        #                     {
        #                         'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  2 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': None,
        #                     },
        #                     {
        #                         'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  2 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': None,
        #                     },
        #                     {
        #                         'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
        #                         'merge_range': {
        #                             "sheetId": sheet_id,
        #                             "startRowIndex":  3 + start_row_index,
        #                             "endRowIndex":  7 + end_row_index,
        #                             "startColumnIndex":  2 + start_col_index,
        #                             "endColumnIndex":   3 + end_col_index
        #                         },
        #                         'merge_type': None,
        #                     }
        #                 ]
        #             },

        #         ]
        #     },

        # ]

        data_body = [
            {
                'team_name': "Team #1",
                'member_tasks': [
                    {
                        'member_name': 'NhuLTQ',
                        'number_of_tasks': 80,
                        'resolve_hour_estimate': {
                            'time': 1,
                            'man_month': 2,
                        },
                        'resolve_hour_actual': {
                            'time': 3,
                            'man_month': 4,
                        },
                        'diff': {
                            'time': 5,
                            'man_month': 6,
                        }
                    },
                    {
                        'member_name': 'NhuLTQ',
                        'number_of_tasks': 80,
                        'resolve_hour_estimate': {
                            'time': 106,
                            'man_month': 106,
                        },
                        'resolve_hour_actual': {
                            'time': 106,
                            'man_month': 106,
                        },
                        'diff': {
                            'time': 106,
                            'man_month': 106,
                        }
                    },
                    {
                        'member_name': 'NhuLTQ',
                        'number_of_tasks': 80,
                        'resolve_hour_estimate': {
                            'time': 106,
                            'man_month': 106,
                        },
                        'resolve_hour_actual': {
                            'time': 106,
                            'man_month': 106,
                        },
                        'diff': {
                            'time': 106,
                            'man_month': 106,
                        }
                    }
                ],
            }
        ]

        def write_and_merge(value):
            range_name = f"'{worksheet_title}'!{chr(start_char + value['merge_range']['startColumnIndex'])}{(value['merge_range']['startRowIndex'] + 1)}"
            if value.get('merge_type', None) is not None:
                self.merge_cells(spreadsheet_id, merge_range=value['merge_range'], merge_type=value['merge_type'])
            else:
                range_name = f"'{worksheet_title}'!{chr(start_char + value['merge_range']['startColumnIndex'])}{(value['merge_range']['startRowIndex'] + 1)}:{
                    chr(start_char + value['merge_range']['startColumnIndex'])}{(value['merge_range']['endRowIndex'])}"
                print('range_name', range_name)
            self.write_data_range(spreadsheet_id, range_name=range_name, data=value['col_name'])

        # create cells header
        for value in cells_header:
            write_and_merge(value)

        # create cell body

        # def handle_cell(values):
        #     i = 0
        #     for value in values:
        #         if i == 1:
        #             values[i]['merge_range']['startRowIndex'] += 5
        #             values[i]['merge_range']['endRowIndex'] += 5
        #         elif i > 1:
        #             values[i]['merge_range']['startRowIndex'] = (values[i-1]['merge_range']['startRowIndex'] + 5)
        #             values[i]['merge_range']['endRowIndex'] = (values[i-1]['merge_range']['endRowIndex'] + 5)
        #         i += 1
        #         write_and_merge(value)

        # for value in cells_body:
        #     # create team
        #     if value.get('teams', None) is not None:
        #         handle_cell(values=value['teams'])

        #     # create members
        #     if value.get('members', None) is not None:
        #         for member in value['members']:
        #             # create total members
        #             if member.get('totals', None) is not None:
        #                 handle_cell(values=member['totals'])
        #             # create team members
        #             if member.get('teams', None) is not None:
        #                 handle_cell(values=member['teams'])
        #             # create names members
        #             if member.get('names', None) is not None:
        #                 handle_cell(values=member['names'])

        # start_row_index = 2
        # end_row_index = 0
        # start_col_index = 0
        # end_col_index = 0
        # for item in cells_body:
        #     pass

        # Step 1: Create a new spreadsheet
           # Create spreadsheet
        spreadsheet_title = "BPO Report"
        spreadsheet = self.create_spreadsheet(spreadsheet_title)
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        # Create worksheet
        worksheet_title = 'Sheet1'
        sheet_id = self.create_worksheet(spreadsheet_id, worksheet_title)

        # Step 2: Prepare the dummy data
        data = [
            ['Team', 'Member', 'Number of tasks', 'Resolve hour estimate', '', '', 'Resolve hour actual', '', '', 'Diff', ''],
            ['Team', 'Member', 'Number of tasks', 'Time', 'Man-month', 'Time', 'Man-month', 'Time', 'Man-month'],
            ['Team #1', '', '', '', '', '', '', '', '', '', ''],
            ['Team', 'Total', 237, 361.2, 3.01, 400.4, 3.34, -39.2, -0.33],
            ['', 'NhuLTQ', 81, 106.4, 0.89, 107.85, 0.90, -1.45, -0.01],
            # ... more dummy data rows as per your screenshot
        ]

        # Step 3: Write data to the spreadsheet
        range_name = f"'Sheet1'!A1:K{len(data)}"
        self.write_data_range(spreadsheet_id, range_name, data)

        # Step 4: Merge cells as per the structure in the screenshot
        merge_ranges = [
            # Merging 'Team' cells vertically
            {
                'sheetId': sheet_id,
                'startRowIndex': 2,
                'endRowIndex': 6,
                'startColumnIndex': 0,
                'endColumnIndex': 1
            },
            # ... more merge ranges as needed
        ]

        for merge_range in merge_ranges:
            self.merge_cells(spreadsheet_id, merge_range)

        # Step 5: Apply formatting
        header_format = {
            "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 1.0},
            "horizontalAlignment": "CENTER",
            "textFormat": {"fontSize": 10, "bold": True}
        }

        # Apply formatting to the header range
        self.format_header(spreadsheet_id, header_format, {
            'sheetId': sheet_id,
            'startRowIndex': 1,
            'endRowIndex': 2,
            'startColumnIndex': 0,
            'endColumnIndex': 10
        })

        print(f"Spreadsheet with ID {spreadsheet_id} created and formatted successfully.")

    def create_table_for_bpo(self, data_json, start_char=65, start_row_index=0, end_row_index=0, start_col_index=0, end_col_index=0):
        # Create spreadsheet
        spreadsheet = self.create_spreadsheet("New Spreadsheet")
        spreadsheet_id = spreadsheet.get('spreadsheetId')

        # # Create worksheet
        # worksheet_title = 'New Worksheet'
        # sheet_id = self.create_worksheet(spreadsheet_id, worksheet_title)
        sheet_id = 0
        merge_requests = []

        # Data
        data = [
            ['Team', 'Member', '', 'Number of tasks', 'Resolve hour estimate', '', 'Resolve hour actual', '', 'Diff'],
            ['', '', '', '', 'Time', 'Man-month', 'Time', 'Man-month', 'Time', 'Man-month'],

        ]

        def generate_merge_request(start_row_index, end_row_index, start_column_index, end_column_index, mergeType='MERGE_ALL'):
            """
                Generate a merge request
            Args:
                start_row_index (int): start row index
                end_row_index (int_): end row index
                start_column_index (int): start column index
                end_column_index (int): end column index
                mergeType (str, optional): Defaults to 'MERGE_ALL'.

            Returns:
                dict : merge request 
            """

            # default merge request
            merge_request = {
                'mergeCells': {
                    'range': {
                        'sheetId': 0,
                        'startRowIndex': start_row_index,
                        'endRowIndex': end_row_index,
                        'startColumnIndex': start_column_index,
                        'endColumnIndex': end_column_index
                    },
                    'mergeType': mergeType
                }
            }
            return merge_request

        def create_merge_request_and_data(data_json):
            """
                Create merge request and data
            Args:
                data_json (list): Data from server
            """

            # Apple merge cell header

            merge_requests.append(generate_merge_request(0, 2, 0, 1))
            merge_requests.append(generate_merge_request(0, 2, 1, 3))
            merge_requests.append(generate_merge_request(0, 2, 3, 4))
            merge_requests.append(generate_merge_request(0, 1, 4, 6))

            merge_requests.append(generate_merge_request(0, 1, 6, 8, mergeType='MERGE_ROWS'))
            merge_requests.append(generate_merge_request(0, 1, 8, 10, mergeType='MERGE_ROWS'))

            # Handle create merge request and data
            row_index = 2
            for item in data_json:
                length_row = item['quantity_member']
                length_col = 0
                list_data = []
                merge_requests.append(generate_merge_request(row_index, row_index + 1 + length_row, length_col, length_col + 1))
                list_data.append(item['team_name'])
                list_data.append('total')
                list_data.extend([''] * 8)
                data.append(list_data)
                list_data = []
                length_col += 1

                for member in item['task_members']:
                    list_data.append('')
                    list_data.append('Team')
                    list_data.append(member['name_member'])
                    list_data.append(member['number_of_tasks'])
                    list_data.append(member['resolve_hour_estimate']['time'])
                    list_data.append(member['resolve_hour_estimate']['man_month'])
                    list_data.append(member['resolve_hour_actual']['time'])
                    list_data.append(member['resolve_hour_actual']['man_month'])
                    list_data.append(member['diff']['time'])
                    list_data.append(member['diff']['man_month'])
                    data.append(list_data)
                    list_data = []
                merge_requests.append(generate_merge_request(row_index, row_index + 1, length_col, length_col + 2, mergeType='MERGE_ROWS'))

                merge_requests.append(generate_merge_request(row_index + 1, row_index + 1 + length_row, length_col, length_col + 1))
                row_index += length_row + 1

        # Call fc to create merge request and create
        create_merge_request_and_data(data_json)

        len_col = len(data[0])
        len_row = len(data)
        merge_requests.append(generate_merge_request(len_row, len_row + 1, 0, 3))
        list_data = []
        list_data.append('total')
        list_data.extend([''] * 7)
        data.append(list_data)
        len_row = len(data)
        range_name = f"{chr(65)}1:{chr(65+len_col)}{len_row}"
        self.merge_cell(spreadsheet_id, requests=merge_requests)
        self.write_data_range(spreadsheet_id, range_name=range_name, data=data)


"""
  - Để merges các cells thì cần phải có các `mergeCells`
  - Nếu mà đã bảng table bộ cố định thì việc định nghĩa các mergeCells rất là dễ
  - Nếu bảng table phụ thuộc vào data để có thể render ra table tương ứng.

---> Vậy câu hỏi đặt ra
  - Em nên dựa vào data để tạo ra các object `mergeCells` và dùng list `object mergeCells` để `format table` trước rồi sau đó mới tìm cách để write data to table đó.
    Ex:
        Ví dụ data có 6 phần tử thì tùy vào đó sẽ tạo ra 6 mergeCell tương ứng và gọi fc để `format table` và sau đó sẽ tìm cách để format data để viết vào table đó.

  - Hay em nên làm 2 việc này cùng một lúc format and write data cùng lúc.
    Ngay sau khi em format em sẽ write data xuống ngay cái ô mà em đã format luôn.
    Ex: 
        Em vẫn tạo `mergeCells` và kem theo data cần ghi vào `mergeCells` đó.
"""
