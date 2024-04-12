import json
import pandas as pd

features = []
excel_file_name = "煙囪.xlsx"

excel_data = pd.read_excel(excel_file_name)

# 迭代CSV檔案中的每一行
for index, row in excel_data.iterrows():
    # 將每一行轉換成需要的JSON格式
    # geometry_dict = json.loads(row["location"].replace("'", '"'))
    feature = {
        "type": "Feature",
        "properties": {
            "Chimney_ID": row["煙囪編號"],
            "Company_Name": row["所屬公司"]
        },
        "geometry":  {
		  "type": "Point",
		  "coordinates": [
			row["經度"],
			row["緯度"]
		  ]
		}
    }
    features.append(feature)

# 將JSON格式寫入檔案
output_json_name = 'output.json'
with open(output_json_name, 'w', encoding='utf-8') as json_file:
    json.dump({"features": features}, json_file, ensure_ascii=False, indent=2)

print(f'Transformation completed. JSON file saved as: {output_json_name}')
