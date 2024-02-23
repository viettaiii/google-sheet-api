
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
                'cells': [
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
                    }
                ]
            },
            {
                'cells': [
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
                    }
                ]
            },
            {
                'cells': [
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
            },

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
            {
                'teams': [
                    {
                        'col_name': [['Team #1']],
                        'merge_range': {
                            "sheetId": sheet_id,
                            "startRowIndex":  2 + start_row_index,
                            "endRowIndex":  7 + end_row_index,
                            "startColumnIndex":  0 + start_col_index,
                            "endColumnIndex":   1 + end_col_index
                        },
                        'merge_type': 'MERGE_ALL',
                    },
                    {
                        'col_name': [['Team #2']],
                        'merge_range': {
                            "sheetId": sheet_id,
                            "startRowIndex":  2 + start_row_index,
                            "endRowIndex":  7 + end_row_index,
                            "startColumnIndex":  0 + start_col_index,
                            "endColumnIndex":   1 + end_col_index
                        },
                        'merge_type': 'MERGE_ALL',
                    },
                    {
                        'col_name': [['Team #3']],
                        'merge_range': {
                            "sheetId": sheet_id,
                            "startRowIndex":  2 + start_row_index,
                            "endRowIndex":  7 + end_row_index,
                            "startColumnIndex":  0 + start_col_index,
                            "endColumnIndex":   1 + end_col_index
                        },
                        'merge_type': 'MERGE_ALL',
                    },
                ]
            },
            {
                'members': [
                    {
                        'totals': [
                            {
                                'col_name': [['Total']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  2 + start_row_index,
                                    "endRowIndex":  3 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                            {
                                'col_name': [['Total']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  2 + start_row_index,
                                    "endRowIndex":  3 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                            {
                                'col_name': [['Total']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  2 + start_row_index,
                                    "endRowIndex":  3 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                        ]
                    },
                    {
                        'teams': [
                            {
                                'col_name': [['Team']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   2 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                            {
                                'col_name': [['Team']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   2 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                            {
                                'col_name': [['Team']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  1 + start_col_index,
                                    "endColumnIndex":   2 + end_col_index
                                },
                                'merge_type': 'MERGE_ALL',
                            },
                        ]
                    },
                    {
                        'names': [
                            {
                                'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  2 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': None,
                            },
                            {
                                'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  2 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': None,
                            },
                            {
                                'col_name': [['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ'], ['NhuLTQ']],
                                'merge_range': {
                                    "sheetId": sheet_id,
                                    "startRowIndex":  3 + start_row_index,
                                    "endRowIndex":  7 + end_row_index,
                                    "startColumnIndex":  2 + start_col_index,
                                    "endColumnIndex":   3 + end_col_index
                                },
                                'merge_type': None,
                            }
                        ]
                    },

                ]
            },

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

        ## create cells header
        for value in cells_header:
            if value.get('cells', None) is not None:
                for cell in value['cells']:
                    write_and_merge(cell)
            else:
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
