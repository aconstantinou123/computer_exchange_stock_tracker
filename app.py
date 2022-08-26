from __future__ import print_function
from enum import Enum
import os.path
import requests
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]
CEX_BASE_URL = "https://wss2.cex.uk.webuy.io/v3"
FILE_NAME = "cex_stock.xlsx"


class Stores(Enum):
    Edinburgh = 54
    Leith = 3115
    CameronToll = 3017


class Categories(Enum):
    DreamCastSoftware = 51
    GamecubeSoftware = 667
    N64Software = 1030
    PlayStation1Software = 1071
    PlayStation2Software = 403
    SuperNintendoSoftware = 1037
    Xbox1Games = 673


def get_stock_data(store_id: str) -> dict:
    category_str = ",".join([str(e.value) for e in Categories])
    games = []
    res = requests.get(
        f"{CEX_BASE_URL}/boxes?storeIds=[{store_id}]&categoryIds=[{category_str}]&firstRecord=0"
    )
    data = res.json()["response"]["data"]
    total_games = data["totalRecords"]
    games.extend(data["boxes"])
    games_retrieved = len(games)
    while games_retrieved < total_games:
        res = requests.get(
            f"{CEX_BASE_URL}/boxes?storeIds=[{store_id}]&categoryIds=[{category_str}]&firstRecord={games_retrieved + 1}"
        )
        data = res.json()["response"]["data"]
        games.extend(data["boxes"])
        games_retrieved = len(games)

    stock_data = {"Category": [], "Title": [], "Price": [], "For Sale": []}
    for game in games:
        stock_data["Category"].append(game["categoryName"])
        stock_data["Title"].append(game["boxName"])
        stock_data["Price"].append(game["sellPrice"])
        stock_data["For Sale"].append(True if game["boxSaleAllowed"] == 1 else False)
    return stock_data


def construct_stock_spreadsheet():
    with pd.ExcelWriter(FILE_NAME) as writer:
        print(f"Generating new stock data for: {', '.join([e.name for e in Stores])}")
        for store in Stores:
            stock_data = get_stock_data(store.value)
            df = pd.DataFrame.from_dict(stock_data)
            df.sort_values(by=["Category", "Title"], inplace=True)
            df.reset_index(drop=True, inplace=True)
            df.to_excel(writer, sheet_name=store.name, index=False)
            worksheet = writer.sheets[store.name]
            for idx, col in enumerate(df):
                series = df[col]
                max_len = (
                    max(
                        (
                            series.astype(str).map(len).max(),
                            len(str(series.name)),
                        )
                    )
                    + 1
                )
                worksheet.set_column(idx, idx, max_len)
        print("New spreadsheet generated")


def google_sign_in():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("creds.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def upload_file(creds):
    try:
        service = build("drive", "v3", credentials=creds)

        results = (
            service.files()
            .list(pageSize=10, q="name='games'", fields="files(name, id)")
            .execute()
        )
        folder_id = results["files"][0]["id"]

        media = MediaFileUpload(
            FILE_NAME,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        folder_id = results["files"][0]["id"]

        existing_files = (
            service.files()
            .list(pageSize=10, q=f"name='{FILE_NAME}'")
            .execute()["files"]
        )
        uploaded_file = None
        if len(existing_files):
            file_metadata = {"name": FILE_NAME}
            file_id = existing_files[0]["id"]
            print(f"{FILE_NAME} exists. Uploading new version")
            uploaded_file = (
                service.files()
                .update(
                    fileId=file_id, body=file_metadata, media_body=media, fields="name"
                )
                .execute()
            )
        else:
            file_metadata = {"name": FILE_NAME, "parents": [folder_id]}
            print(f"Uploading {FILE_NAME}")
            uploaded_file = (
                service.files()
                .create(body=file_metadata, media_body=media, fields="name")
                .execute()
            )
        print(f'File uploaded: {uploaded_file.get("name")}')

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    construct_stock_spreadsheet()
    creds = google_sign_in()
    upload_file(creds)
