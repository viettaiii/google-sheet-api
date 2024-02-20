
from handle_data import GoogleSheetManager
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", 'https://www.googleapis.com/auth/spreadsheets']


def main():
    creds_file = "credentials.json"
    sheet_api = GoogleSheetManager(creds_file, SCOPES)
    sheet_api.build_service("sheets", "v4")
    spreadsheet = sheet_api.connect_spreadsheet('1lX4oJi9KliX_G3-UTPOQaUzydMSXhqcCeS6YU8ZpGV8')
    # spreadsheet_id = spreadsheet.get('spreadsheetId')
    # print(sheet_api.get_worksheet_by_title(spreadsheet.get('spreadsheetId'), 'Sheet1'))
    # sheet_api.create_worksheet(spreadsheet.get('spreadsheetId'), 'worksheet 2')
    # sheet_api.create_spreadsheet("")
    data = [
        ["STT",  'Họ', 'Tên', 'SDT', 'Ngày sinh'],
        [1, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [2, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [3, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [4, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [5, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [6, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    ]
    # sheet_api.write_data_range(spreadsheet.get('spreadsheetId'), "'Sheet2'!A1:B2", data)
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
        "sheetId": None,
        "startRowIndex": 0,
        "endRowIndex": len(data),
        "startColumnIndex": 0,
        "endColumnIndex": 0
    }

    sheet_api.import_data_to_google_sheet(data, format_cell, format_range,merge_range)



main()
