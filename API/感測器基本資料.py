#
# 查詢環保署智慧城鄉空品微型感測器所有感測器基本資料以及位置(Location)資訊
# SensorThings API: https://sta.colife.org.tw/STA_AirQuality_EPAIoT/v1.0/Things?$expand=Locations
#

import requests
import json
import urllib3
import datetime
import pandas as pd


urllib3.disable_warnings()

print("begin time: ", datetime.datetime.now())

# 取得 API 的回應
response = requests.get("https://sta.colife.org.tw/STA_AirQuality_EPAIoT/v1.0/Things?$expand=Locations",verify=False)

# 解析 JSON 格式的回應
data = json.loads(response.content)

# 取得資料總筆數
count = data["@iot.count"]
# print(count)

m = 0
data_dict = []

# 取得前100筆資料
for item in data["value"]:
    name = item["name"]
    properties = item["properties"]
    stationID = properties["stationID"]

    for item1 in item["Locations"]:
        location = item1["location"]

        if location["coordinates"][1]<25.123979 and location["coordinates"][1]>24.587723 and location["coordinates"][0]>120.98353 and location["coordinates"][0]<121.48448:
            m = m+1
            my_dict = {}
            my_dict["stationID"] = stationID
            my_dict["location"] = location
            data_dict.append(my_dict)

            print(str(m), " : ", name, " , " , stationID, " , " , location["coordinates"])

# print("page 1: ", m, " time: ", datetime.datetime.now())

# 計算資料分頁數
count1 = int(count/100)+1
# print(count1)

# 檢查是否有 @iot.nextLink 參數，若有代表有分頁
str_link = data["@iot.nextLink"]

# 檢查是否有 @iot.nextLink 參數，若有代表有分頁
if (not("@iot.nextLink" in data)):
    # 無超過一頁之資料
    print("")
else:
    #print("No it is not empty")
    # 取得各分頁資料
    del item, item1, data

    # 下面迴圈寫法會列出所有分頁感測器：
    # for i in range(100, count1*100, 100):
    # 此範例只列部份感測器資料：
    for i in range(100, count1*100, 100):
        url = "https://sta.colife.org.tw/STA_AirQuality_EPAIoT/v1.0/Things?$expand=Locations&$skip="+str(i)

        response = requests.get(url,verify=False)

        data = json.loads(response.content)

        for item in data["value"]:
            name = item["name"]
            properties = item["properties"]
            stationID = properties["stationID"]

            for item1 in item["Locations"]:
                location = item1["location"]

                if location["coordinates"][1]<25.123979 and location["coordinates"][1]>24.587723 and location["coordinates"][0]>120.98353 and location["coordinates"][0]<121.48448:
                    m = m+1
                    my_dict = {}
                    my_dict["stationID"] = stationID
                    my_dict["location"] = location
                    data_dict.append(my_dict)

                    print(str(m), " : ", name, " , " , stationID, " , " , location["coordinates"])
        del item, item1, data
        

print("total : ", m)



# 將字典轉換為 DataFrame
df = pd.DataFrame(data_dict)

# 將 DataFrame 寫入 Excel 檔案
excel_file_path = 'output.xlsx'
df.to_excel(excel_file_path, index=False)

print(f'DataFrame 已寫入到 {excel_file_path}')