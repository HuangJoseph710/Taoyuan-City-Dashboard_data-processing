import pandas as pd
import json
from datetime import datetime, timezone, timedelta
import numpy as np


excel_file_path = 'AQI歷史資料.xlsx'

excel_data = pd.read_excel(excel_file_path)


data_dict = excel_data.to_dict(orient='records')

# for i in range(7):
#     location['sitename']
#     print(data_dict[i]['longitude'])
#     print()

data1 = []
data2 = []
data3 = []
data4 = []
data5 = []
data6 = []

timezone_info = timezone(timedelta(hours=8))
for i in range(0,len(data_dict),6):
    
    my_dict = {}
    my_dict['y'] = (np.nan_to_num(data_dict[i]["pm2.5"], nan=data_dict[i]["pm2.5_avg"]))
    my_dict['x'] = pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data1.append(my_dict)
    
    my_dict = {}
    my_dict['y'] = (np.nan_to_num(data_dict[i+1]["pm2.5"], nan=data_dict[i+1]["pm2.5_avg"]))
    my_dict['x'] = pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data2.append(my_dict)
    
    my_dict = {}
    my_dict['y'] = (np.nan_to_num(data_dict[i+2]["pm2.5"], nan=data_dict[i+2]["pm2.5_avg"]))
    my_dict['x'] = pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data3.append(my_dict)
    
    my_dict = {}
    my_dict['y'] = (np.nan_to_num(data_dict[i+3]["pm2.5"], nan=data_dict[i+3]["pm2.5_avg"]))
    my_dict['x'] = pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data4.append(my_dict)
    
    my_dict = {}
    
    my_dict['y'] = (np.nan_to_num(data_dict[i+4]["pm2.5"], nan=data_dict[i+4]["pm2.5_avg"]))
    my_dict['x'] =pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data5.append(my_dict)
    
    my_dict = {}
    my_dict['y'] = (np.nan_to_num(data_dict[i+5]["pm2.5"], nan=data_dict[i+5]["pm2.5_avg"]))
    my_dict['x'] = pd.Timestamp(data_dict[i]["datacreationdate"]).strftime('%Y-%m-%dT%H:%M:%S%z+08:00')
    data6.append(my_dict)

data1.reverse()
data2.reverse()
data3.reverse()
data4.reverse()
data5.reverse()
data6.reverse()


final_data = {
	"data": [
		{
			"name": "中壢",
			"data": data1
		},
		{
			"name": "龍潭",
			"data": data2
		},
		{
			"name": "平鎮",
			"data": data3
		},
		{
			"name": "觀音",
			"data": data4
		},
		{
			"name": "大園",
			"data": data5
		},
		{
			"name": "桃園",
			"data": data6
		}
	]
}

# file_path = '203.json'
# with open(file_path, 'w', encoding='utf-8') as json_file:
#     json.dump(final_data, json_file, indent=4, ensure_ascii=False)

# print(f'JSON 文件已寫入到 {file_path}')





# 創建一個 ExcelWriter 對象，指定 Excel 檔案路徑
excel_file_path = 'output_with_multiple_sheets.xlsx'
with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
    # 寫入每個工作表
    pd.DataFrame(data1).to_excel(writer, sheet_name='Sheet1', index=False)
    pd.DataFrame(data2).to_excel(writer, sheet_name='Sheet2', index=False)
    pd.DataFrame(data3).to_excel(writer, sheet_name='Sheet3', index=False)
    pd.DataFrame(data4).to_excel(writer, sheet_name='Sheet4', index=False)
    pd.DataFrame(data5).to_excel(writer, sheet_name='Sheet5', index=False)
    pd.DataFrame(data6).to_excel(writer, sheet_name='Sheet6', index=False)

print(f'DataFrames 已寫入到 {excel_file_path}')