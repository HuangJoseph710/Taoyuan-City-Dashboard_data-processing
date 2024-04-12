import pandas as pd
import json
from datetime import datetime, timezone, timedelta
import numpy as np


excel_file_path = '異常發生原因.xlsx'

excel_data = pd.read_excel(excel_file_path)


data_dict = excel_data.to_dict(orient='dict')

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






final_data = {
	"data": [
		{
			"name": "火災",
			"data": list(data_dict["火災"].values())
		},
		{
			"name": "工廠排放廢氣",
			"data": list(data_dict["工廠排放廢氣"].values())
		},
		{
			"name": "焚化爐",
			"data": list(data_dict["焚化爐"].values())
		},
		{
			"name": "其他人為因素",
			"data": list(data_dict["其他人為因素"].values())
		},
		{
			"name": "未知原因",
			"data": list(data_dict["未知原因"].values())
		},
		{
			"name": "總計",
			"data": list(data_dict["總計"].values())
		}
	]
}

file_path = '206.json'
with open(file_path, 'w', encoding='utf-8') as json_file:
    json.dump(final_data, json_file, indent=4, ensure_ascii=False)

print(f'JSON 文件已寫入到 {file_path}')





# # 創建一個 ExcelWriter 對象，指定 Excel 檔案路徑
# excel_file_path = 'output_with_multiple_sheets.xlsx'
# with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
#     # 寫入每個工作表
#     pd.DataFrame(data1).to_excel(writer, sheet_name='Sheet1', index=False)
#     pd.DataFrame(data2).to_excel(writer, sheet_name='Sheet2', index=False)
#     pd.DataFrame(data3).to_excel(writer, sheet_name='Sheet3', index=False)
#     pd.DataFrame(data4).to_excel(writer, sheet_name='Sheet4', index=False)
#     pd.DataFrame(data5).to_excel(writer, sheet_name='Sheet5', index=False)
#     pd.DataFrame(data6).to_excel(writer, sheet_name='Sheet6', index=False)

# print(f'DataFrames 已寫入到 {excel_file_path}')