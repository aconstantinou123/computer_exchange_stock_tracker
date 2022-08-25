from enum import Enum
import requests
import pandas as pd


CEX_BASE_URL = "https://wss2.cex.uk.webuy.io/v3"


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
    with pd.ExcelWriter("cex.xlsx") as writer:
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


construct_stock_spreadsheet()
