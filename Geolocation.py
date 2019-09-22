import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import pandas as pd
import googlemaps
import gmap_key

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("Cred.json", scope)
client = gspread.authorize(creds)
sheet = client.open("GeolocationDemo").sheet1 #Change here
gmaps_key = googlemaps.Client(key = gmap_key.api_key) #Change here

row = sheet.row_values(1)
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

col = sheet.col_values(1)
for i in range(2, len(col)+1):
    row = sheet.row_values(i)
    geocode_result = gmaps_key.geocode("{}, {}, Udaipur, India".format(row[adr_index-1], row[area_index-1]))
    try:
        lat = geocode_result[0]["geometry"]["location"]["lat"]
        lng = geocode_result[0]["geometry"]["location"]["lng"]
    except:
        lat = None
        lng = None
    sheet.update_cell(i, lat_index, lat)
    sheet.update_cell(i, long_index, lng)