# Google Sheet API

## Mục lục
* [Chuẩn bị môi trường](#chuan-bi)
* [Cài đặt thư viện](#cai-dat-thu-vien)
* [Cài đặt vào dự án](#cai-dat-vao-du-an)
* [Các hàm tiện ích](#cac-ham-tien-ich)
* [Tài liệu thảm khảo](#tai-lieu-tham-khao)

## Chuẩn bị môi trường
- Vào link nay để setup theo hướng dẫn: [Chuẩn bị môi trường](https://developers.google.com/sheets/api/quickstart/python?hl=vi#set_up_your_environment)
- Sau khi có được file `credentials.json` đến bước `Cài đặt thư viện`.

## Cài đặt thư viện
- Sử dụng câu lệnh sau đề cài đặt để sử dụng thư viện:
    ```python
    # Đối với pip
    pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

    # Đối với pip3
    pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
    ```
## Cài đặt vào dự án
```python
# Code example

from handle_data import GoogleSheetManager

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly", 'https://www.googleapis.com/auth/spreadsheets']

def main():
    # file credentials.json đã setup trước đó.
    creds_file = "credentials.json"

    # Khởi tạo một google sheet manager
    sheet_api = GoogleSheetManager(creds_file, SCOPES)

    # Build service và xác thực người dùng.
    sheet_api.build_service("sheets", "v4")
main()
```
> Sau khi chạy xong thì nó sẽ tạo ra một file `token.json`

## Các hàm tiện ích
- Các hàm tiện ích bằng cách sử dụng sheet_api.[Name Function]

| Name Function | Parameters | Description | Return |
|----------|----------|----------|----------|
| import_data_to_google_sheet   |  data(list[list])<br/> format_cel(dict)<br/> format_range(dict)| `Import` dữ liệu đến google drive  | None | 
| create_spreadsheet | title (str) | Tạo mới một spreadsheet đến google sheet | Spreadsheet |
| connect_spreadsheet | spreadsheet_id (str) | Kết nối đến Spreadsheet | Spreadsheet|
| get_worksheet_by_title |spreadsheet_id(str) <br> title(str) | get worksheet bằng title | Worksheet|
| get_values_range |spreadsheet_id(str) <br> range_name(dict) | Get values dựa trên range | List|
| write_data_range |spreadsheet_id(str) <br> range_name(dict)<br>data(list[list]) | Write data to google sheet | None|
| format_header | spreadsheet_id(str)<br>  cell_format(dict), format_range(dict) | Format cell | None|
| merge_cells | spreadsheet_id(str)<br>  merge_range(dict)<br/> merge_type(str) | Format cell | None|

>>> Chi tiết về format_range, format_cell tham khảo link: [Format](https://developers.google.com/sheets/api/samples/formatting?hl=vi)

```python
    # Code Example

    # Dữ liệu
    data = [
        ["STT",  'Họ', 'Tên', 'SDT', 'Ngày sinh'],
        [1, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [2, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [3, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [4, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [5, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
        [6, "Nguyễn Viết", 'Tài', '0329638260', '2003/02/21'],
    ]

    # Example format_cell, tùy vào tình huống ta có thể config `format_cell` theo định dạng mà ta muốn.

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

    # Phạm vi format cell
    format_range = {
        "sheetId": None, # worksheet id
        "startRowIndex": 0, # Start at A1 in excel
        "endRowIndex": 1,  
        "startColumnIndex": 0, # Start at A1 in excel
        "endColumnIndex":len(data[0])
    }

    # Thực hiện import dữ liệu đến google sheet
    sheet_api.import_data_to_google_sheet(data, format_cell, format_range)

```

## Tài liệu thảm khảo
- https://developers.google.com/sheets/api/samples/formatting?hl=vi
- https://developers.google.com/sheets/api/quickstart/python?hl=vi
- https://github.com/viettaiii/google-sheet-api