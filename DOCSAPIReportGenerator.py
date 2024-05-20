from googleapiclient.discovery import build
import os.path
from google.oauth2.credentials import Credentials
from datetime import datetime

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/drive.file",
]

DOCUMENT_ID = "1XTePwmJewSp06yXK8MbolfDVRIvBFLEAqLAG9qQrrdg"


def main(new_id):
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    service = build("docs", "v1", credentials=creds)

    get_time = datetime.today().strftime('%m/%d/%Y')
    date = get_time
    sales = input("Net sales: ")
    labor = input("Labor cost: ")
    today_labor_percent = input("Enter today labor %: ")
    today_diff = input("Today Up/Down % (whole number): ")

    print("===== Hourly Sales =====")
    sale11 = input("11:00 AM sales: ")
    sale12 = input("12:00 PM sales: ")
    sale13 = input("1:00 PM sales: ")
    sale14 = input("2:00 PM sales: ")
    sale15 = input("3:00 PM sales: ")

    first_half_sales = float(sale11) + float(sale12) + float(sale13) + float(sale14) + float(sale15)
    second_half_sales = float(sales) - float(first_half_sales)

    print("===== Last Week Input ====")
    last_week_sales = input("Enter last week sales: ")
    last_week_labor_percent = input("Enter last week labor %: ")
    last_week_diff = input("Last Week Up/Down % (whole number): ")

    requests = [
        {
            'insertText': {
                'location': {
                    'index': 231,
                },
                'text': last_week_diff + " "
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 226,
                },
                'text': today_diff + " "
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 223,
                },
                'text': '{:,.2f}'.format(float(last_week_labor_percent))
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 209,
                },
                'text': '{:,.2f}'.format(float(today_labor_percent))
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 174,
                },
                'text': '{:,.2f}'.format(float(last_week_sales))
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 170,
                },
                'text': '{:,.2f}'.format(float(sales))
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 166,
                },
                'text': '{:,.2f}'.format(second_half_sales)
            }
        },
        {
            'insertText': {
                'location': {
                    'index': 162,
                },
                'text': '{:,.2f}'.format(first_half_sales)
            }
        },
        # Store Overview: Labor
        {
            'insertText': {
                'location': {
                    'index': 87,
                },
                'text': '{:,.2f}'.format(float(labor))
            }
        },
        # Store Overview: Sales
        {
            'insertText': {
                'location': {
                    'index': 75,
                },
                'text': '{:,.2f}'.format(float(sales))
            }
        },
        # Store Overview: Date
        {
            'insertText': {
                # index 위치가 43번인 곳에 insertText
                'location': {
                    'index': 43,
                },
                'text': date
            }
        },
    ]

    result = service.documents().batchUpdate(
        documentId=new_id, body={'requests': requests}).execute()
    print('updated : {0}'.format(result))


def copy_original():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # Using Drive API
    drive_service = build("drive", "v3", credentials=creds)

    # request body
    get_time = datetime.today().strftime('%m/%d/%Y')
    copy_title = get_time + " Seattle Daily Report"
    body = {
        'name': copy_title
    }

    # copy
    drive_response = drive_service.files().copy(fileId=DOCUMENT_ID, body=body).execute()
    document_copy_id = drive_response.get('id')

    # print result
    print('Copied document with document_id : {0}'.format(document_copy_id))
    return document_copy_id


if __name__ == "__main__":
    new_doc_id = copy_original()
    main(new_doc_id)

    input("close terminal")
