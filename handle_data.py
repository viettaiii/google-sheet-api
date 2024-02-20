
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
