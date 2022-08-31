from __future__ import print_function
import os
from enum import Enum
from datetime import datetime
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
EXISTING_FILE = "existing_cex_stock.xlsx"

DATE_FMT = "%d/%m/%Y"


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


class Status(Enum):
    New = "NEW"
    InStock = "-"
    Sold = "SOLD"


class Columns(Enum):
    Category = "Category"
    Title = "Title"
    Price = "Price"
    ForSale = "For Sale"
    Status = "Status"
    DateAddedOrRemoved = "Date Added/Removed"


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

    stock_data = {
        Columns.Category.value: [],
        Columns.Title.value: [],
        Columns.Price.value: [],
        Columns.ForSale.value: [],
    }
    for game in games:
        stock_data[Columns.Category.value].append(game["categoryName"])
        stock_data[Columns.Title.value].append(game["boxName"])
        stock_data[Columns.Price.value].append(game["sellPrice"])
        stock_data[Columns.ForSale.value].append(
            True if game["boxSaleAllowed"] == 1 else False
        )
    return stock_data


def compare_existing_stock(store, new_stock: dict) -> dict:
    workbook = pd.ExcelFile(EXISTING_FILE)
    existing_df = pd.read_excel(workbook, sheet_name=store)
    existing_stock = dict(existing_df)

    all_stock = {
        Columns.Category.value: [],
        Columns.Title.value: [],
        Columns.Price.value: [],
        Columns.ForSale.value: [],
        Columns.Status.value: [],
        Columns.DateAddedOrRemoved.value: [],
    }

    all_stock = format_new_stock(existing_stock, new_stock, all_stock)
    all_stock = format_existing_stock(existing_stock, new_stock, all_stock)
    return remove_sold_stock(all_stock)


def format_new_stock(existing_stock: dict, new_stock: dict, all_stock: dict) -> dict:
    for index, title in enumerate(list(new_stock[Columns.Title.value])):
        all_stock[Columns.Category.value].append(
            new_stock[Columns.Category.value][index]
        )
        all_stock[Columns.Title.value].append(new_stock[Columns.Title.value][index])
        all_stock[Columns.Price.value].append(new_stock[Columns.Price.value][index])
        all_stock[Columns.ForSale.value].append(new_stock[Columns.ForSale.value][index])
        if title not in list(existing_stock[Columns.Title.value]):
            today_str = datetime.today().strftime(DATE_FMT)
            all_stock[Columns.Status.value].append(Status.New.value)
            all_stock[Columns.DateAddedOrRemoved.value].append(today_str)
        elif stock_age(title, existing_stock) <= 1:
            existing_date = find_existing_date(title, existing_stock)
            all_stock[Columns.Status.value].append(Status.New.value)
            all_stock[Columns.DateAddedOrRemoved.value].append(existing_date)
        else:
            existing_date = find_existing_date(title, existing_stock)
            all_stock[Columns.Status.value].append(Status.InStock.value)
            all_stock[Columns.DateAddedOrRemoved.value].append(existing_date)
    return all_stock


def format_existing_stock(
    existing_stock: dict, new_stock: dict, all_stock: dict
) -> dict:
    for index, title in enumerate(existing_stock[Columns.Title.value]):
        if title not in new_stock[Columns.Title.value]:
            today_str = datetime.today().strftime(DATE_FMT)
            all_stock[Columns.Category.value].append(
                existing_stock[Columns.Category.value][index]
            )
            all_stock[Columns.Title.value].append(
                existing_stock[Columns.Title.value][index]
            )
            all_stock[Columns.Price.value].append(
                existing_stock[Columns.Price.value][index]
            )
            all_stock[Columns.ForSale.value].append(
                existing_stock[Columns.ForSale.value][index]
            )
            if existing_stock[Columns.Status.value][index] != Status.Sold.value:
                all_stock[Columns.Status.value].append(Status.Sold.value)
                all_stock[Columns.DateAddedOrRemoved.value].append(today_str)
            else:
                all_stock[Columns.Status.value].append(
                    existing_stock[Columns.Status.value][index]
                )
                all_stock[Columns.DateAddedOrRemoved.value].append(
                    existing_stock[Columns.DateAddedOrRemoved.value][index]
                )
    return all_stock


def remove_sold_stock(all_stock: dict) -> dict:
    today = datetime.today()
    filtered_stock = {
        Columns.Category.value: [],
        Columns.Title.value: [],
        Columns.Price.value: [],
        Columns.ForSale.value: [],
        Columns.Status.value: [],
        Columns.DateAddedOrRemoved.value: [],
    }
    for index, status in enumerate(all_stock[Columns.Status.value]):
        existing_date = datetime.strptime(
            all_stock[Columns.DateAddedOrRemoved.value][index], DATE_FMT
        )
        stock_age = (today - existing_date).days
        if status != Status.Sold.value or (
            status == Status.Sold.value and stock_age <= 1
        ):
            for k, v in filtered_stock.items():
                filtered_stock[k].append(all_stock[k][index])
    return filtered_stock


def stock_age(title: str, existing_stock: dict) -> int:
    existing_date_str = find_existing_date(title, existing_stock)
    today = datetime.today()
    existing_date = datetime.strptime(existing_date_str, DATE_FMT)
    delta = today - existing_date
    return delta.days


def find_existing_date(title: str, existing_stock: dict) -> str:
    for index, existing_title in enumerate(existing_stock[Columns.Title.value]):
        if title == existing_title:
            return existing_stock[Columns.DateAddedOrRemoved.value][index]


def highlight_cells(value):
    color = "white"
    if value.Status == Status.New.value:
        color = "green"
    elif value.Status == Status.Sold.value:
        color = "red"
    return [f"background-color: {color}"] * len(value)


def construct_stock_spreadsheet():
    os.rename(FILE_NAME, EXISTING_FILE)
    with pd.ExcelWriter(FILE_NAME) as writer:
        print(f"Generating new stock data for: {', '.join([e.name for e in Stores])}")
        for store in Stores:
            print(f"{store.name} stock data generated")
            stock_data = get_stock_data(store.value)
            stock_data_with_status = compare_existing_stock(store.name, stock_data)
            df = pd.DataFrame.from_dict(stock_data_with_status)
            df.sort_values(
                by=[Columns.Category.value, Columns.Title.value], inplace=True
            )
            df.reset_index(drop=True, inplace=True)
            df_styler = df.style.apply(highlight_cells, axis=1)
            df_styler.to_excel(writer, sheet_name=store.name, index=False)

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
