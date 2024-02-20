from googleapiclient.errors import HttpError
from authenticate import AuthenticateGoogleApi
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleSheetApi(AuthenticateGoogleApi):
    def __init__(self, creds_file, scopes) -> None:
        super().__init__(creds_file, scopes)

    def build_service(self, api_name, api_version):
        """
            Create a Google sheet service
        Args:
            api_name (str): Name api 
            api_version (str): version
        Returns: service
        """
        try:
            if self.creds is None:
                self.authenticate()
            self.service = build(api_name, api_version, credentials=self.creds)
            return self.service
        except HttpError as err:
            print(err)

    def create_spreadsheet(self, title):
        """
            Create new spreadsheet

        Args:
            title (string): title of the spreadsheet

        Returns: 
            - Spreadsheet: spreadsheet
        """
        try:
            spreadsheet = {"properties": {"title": title}}
            spreadsheet = (
                self.service.spreadsheets()
                .create(body=spreadsheet, fields="spreadsheetId")
                .execute()
            )
            print(f"Spreadsheet ID: {(spreadsheet.get('spreadsheetId'))}")
            return spreadsheet
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def connect_spreadsheet(self, spreadsheet_id):
        """
            Connect spreadsheet

        Args:
            spreadsheet_id (string): spreadsheet id

        Returns: 
            - Spreadsheet: spreadsheet
        """
        try:
            spreadsheet = (
                self.service.spreadsheets()
                .get(spreadsheetId=spreadsheet_id)
                .execute()
            )
            return spreadsheet
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def get_worksheet_by_title(self, spreadsheet_id, title):
        """
        Connect to a worksheet in the specified spreadsheet based on its title.

        Args:
            spreadsheet_id (str): The ID of the spreadsheet.
            title (str): The title of the worksheet.

        Returns:
            Worksheet or None: Information about the worksheet if found, None otherwise.
        """
        try:
            # Lấy thông tin về tất cả các worksheet trong spreadsheet
            spreadsheet = (
                self.service.spreadsheets()
                .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
                .execute()
            )
            sheets = spreadsheet.get("sheets", [])

            # Duyệt qua danh sách các worksheet và tìm worksheet với tiêu đề mong muốn
            for sheet in sheets:
                properties = sheet.get("properties", {})
                if properties.get("title") == title:
                    return sheet

            # Nếu không tìm thấy worksheet, trả về None
            print(f"Worksheet '{title}' not found.")
            return None
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    def get_values_range(self, spreadsheet_id, range_name):
        """
            Get values of spreadsheet on range

        Args:
            spreadsheet_id (string): spreadsheet id
            range_name (string): range

        Returns: 
            list : list records
        """
        try:
            result = (
                self.service.spreadsheets()
                .values()
                .get(spreadsheetId=spreadsheet_id, range=range_name)
                .execute()
            )
            rows = result.get("values", [])
            print(f"{len(rows)} rows retrieved")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    def create_worksheet(self, spreadsheet_id, title):
        """
            Create a new worksheet
        Args:
            spreadsheet_id (_type_): _description_
            title (_type_): _description_
        Returns:
            str : worksheet id
        """
        try:
            # Tạo request để thêm một sheet mới vào spreadsheet
            request = {
                "addSheet": {
                    "properties": {
                        "title": title
                    }
                }
            }
            # Thực hiện request
            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={'requests': [request]}).execute()
            # Trả về ID của worksheet mới tạo
            return response["replies"][0]["addSheet"]["properties"]["sheetId"]
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    # write data to worksheet
    def write_data_range(self, spreadsheet_id, range_name, data):
        try:
            body = {"values": data}
            result = (
                self.service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption="USER_ENTERED",
                    body=body,
                )
                .execute()
            )
            print(f"{result.get('updatedCells')} cells updated.")
            return result
        except HttpError as error:
            print(f"An error occurred: {error}")
            return error

    # format header
    def format_header(self, spreadsheet_id,  cell_format=None, format_range={}):
        try:
            # Tạo request để định dạng ô
            request = {
                "repeatCell": {
                    "range": format_range,
                    "cell": {
                        "userEnteredFormat": cell_format
                    },
                    "fields": "userEnteredFormat"
                }
            }

            # Thực hiện request
            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={'requests': [request]}).execute()

            # In thông báo và trả về kết quả
            print("Cell formatted successfully.")
            return response
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    # merge cell

    def merge_cells(self, spreadsheet_id, merge_range, merge_type='MERGE_ROWS'):
        """
            Merge cell
        Args:
            spreadsheet_id (str): spreadsheet id
            merge_range (dict): dictionary range to merge cell
        """
        try:
            requests = {
                "mergeCells": {
                    "range": merge_range,
                    "mergeType": merge_type
                }
            },
            response = self.service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={'requests': [requests]}).execute()
            return None
        except HttpError as error:
            print(f"An error occurred: {error}")
            return None

    def merge_cell_and_write_data(self, spreadsheet_id, merge_range,  range_name, data, merge_type=None):
        for row in data:
            for i in range(0, len(row)):
                if i == merge_range['startColumnIndex']:
                    for j in range(2, merge_range['endColumnIndex']):
                        row[i] = row[i] + ' ' + row[j]
                    break
        self.merge_cells(spreadsheet_id, merge_range)
        self.write_data_range(spreadsheet_id, range_name, data)
