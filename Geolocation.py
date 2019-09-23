import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import pandas as pd
import googlemaps
import gmap_key

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("Cred.json", scope)
client = gspread.authorize(creds)
sheet = client.open("GeolocationDemo")
sheetId = sheet.worksheet("Sheet1")
gmaps_key = googlemaps.Client(key = gmap_key.api_key)

row = sheetId.row_values(1)
adr_index = 1
area_index = 1
lat_index = 1
long_index = 1

for i in row:
    if i == "Address":
        break
    adr_index += 1
for i in row:
    if i == "Area":
        break
    area_index += 1
for i in row:
    if i == "Latitude":
        break
    lat_index += 1
for i in row:
    if i == "Longitude":
        break
    long_index += 1

col = sheetId.col_values(1)
for i in range(2, len(col)+1):
    row = sheetId.row_values(i)
    geocode_result = gmaps_key.geocode("{}, {}, Udaipur, India".format(row[adr_index-1], row[area_index-1]))
    try:
        lat = geocode_result[0]["geometry"]["location"]["lat"]
        lng = geocode_result[0]["geometry"]["location"]["lng"]
    except:
        lat = None
        lng = None
    sheetId.update_cell(i, lat_index, lat)
    sheetId.update_cell(i, long_index, lng)

reqs = {
    "requests": [
        {
            "repeatCell": {
                "range": {
                    "endRowIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 1.0,
                        "green": 8.0,
                        "blue": 0.0
                    },
                    "horizontalAlignment" : "CENTER",
                    "textFormat": {
                        "foregroundColor": {
                        "red": 0.0,
                        "green": 0.0,
                        "blue": 0.0
                        },
                        "fontSize": 12,
                    }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            },
            "setDataValidation": {
                "range": {
                    "startRowIndex": 1,
                    "endRowIndex": len(col),
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                },
                "rule": {
                    "condition": {
                        "type": "ONE_OF_LIST",
                        "values": [
                            {"userEnteredValue": "PENDING"},
                            {"userEnteredValue": "VISITED"}
                        ]
                    },
                    "showCustomUi": True
                }
            }
        }
    ]
}
sheet.batch_update(reqs)