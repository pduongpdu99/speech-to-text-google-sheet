from lib.Google import Create_Service
from googleapiclient.http import MediaFileUpload
from localStoragePy import localStoragePy
import os

CLIENT_SECRET_FILE = 'client_secret_GoogleCloudDemo.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']


def local_storage_clear(appname: str, backend: str):
    localStorage = localStoragePy(appname, backend)
    localStorage.clear()


def service_init(version):
    return Create_Service(CLIENT_SECRET_FILE, API_NAME, version, SCOPES)


def create_response(service, media, file_metadata):
    return service.files().create(
        media_body=media,
        body=file_metadata,
    ).execute()


def update_response(service, fileId, media, file_metadata):
    return service.files().update(
        fileId=fileId,
        media_body=media,
        body=file_metadata,
    ).execute()


def convert_excel_file(filepath: str, folder_ids: list = None, fileId: int = None):
    if not os.path.exists(filepath):
        print(f"{filepath} not found")
        return

    localStorage = localStoragePy('excel-to-googlesheet', 'drive')

    try:
        file_metadata = {
            'name': os.path.splitext(os.path.basename(filepath))[0],
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': folder_ids
        }

        media = MediaFileUpload(
            filepath,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

        fileId = localStorage.getItem('fileId')
        if fileId != None:
            service = service_init("v2")
            response = update_response(
                service,
                fileId=fileId,
                media_body=media,
                body=file_metadata
            )
            print("Complete updated")

        else:
            service = service_init(API_VERSION)
            response = create_response(
                service,
                media_body=media,
                body=file_metadata
            )

            localStorage.setItem('fileId', response['id'])
            print("Complete created")

        return response
    except Exception as e:
        print(f"Error {str(e)}")


def export_program():
    # excel_files = os.listdir('./')
    convert_excel_file(os.path.join("./", 'test.xlsx'),
                       ['1VcxrWjvLhXrlT0exUJDliBrtjQh9yVuN'])

    # for excel_file in excel_files:
    #     if ".xlsx" in excel_file and "~$" not in excel_file:
    #         convert_excel_file(os.path.join("./", excel_file),
    #                         ['1VcxrWjvLhXrlT0exUJDliBrtjQh9yVuN'])
