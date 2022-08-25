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

category_str = ",".join([str(e.value) for e in Categories])

games = []
res = requests.get(f"{CEX_BASE_URL}/boxes?storeIds=[{Stores.Edinburgh.value}]&categoryIds=[{category_str}]&firstRecord=0")
data = res.json()['response']['data']
total_games = data['totalRecords']
games.extend(data['boxes'])
games_retrieved = len(games)
while games_retrieved < total_games:
    res = requests.get(f"{CEX_BASE_URL}/boxes?storeIds=[{Stores.Edinburgh.value}]&categoryIds=[{category_str}]&firstRecord={games_retrieved + 1}")
    data = res.json()['response']['data']
    games.extend(data['boxes'])
    games_retrieved = len(games)


store_data = {
    'Category': [],
    'Title': [],
    'Price': [],
    'For Sale': []

}
for game in games:
    store_data['Category'].append(game['categoryName'])
    store_data['Title'].append(game['boxName'])
    store_data['Price'].append(game['sellPrice'])
    store_data['For Sale'].append(True if game['boxSaleAllowed'] == 1 else False)

df = pd.DataFrame.from_dict(store_data)
df.sort_values(by=['Category', 'Title'], inplace=True)
df.reset_index(drop=True, inplace=True)
with pd.ExcelWriter("cex.xlsx") as writer:
    df.to_excel(writer, sheet_name="Edinburgh", index=False)  

