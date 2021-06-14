import pandas as pd
import json

# Change name file input
xlsx = pd.ExcelFile("data/test.xlsx")
df = pd.read_excel(xlsx, "data")

# Get another column
id = df['id']
day = df['day']
time = df['time']
mau_bantudong_id = df['sampleid']
purpose_id = df['note']
tnn_station_id = df['stationid']
qualityindex_id = df['qualityindexid']

# Get detail column
detail = df['detail']
nrows = len(df)

# Create Empty Array to Export Excel
std_para = []
value_arr = []
inlimit_arr = []

id_arr = []
date = []
mau_bantudong_id_arr = []
purpose_id_arr = []
tnn_station_id_arr = []
qualityindex_id_arr = []

# Main App
for row in range(0, nrows):
    detail_row = detail[row]
    # Convert JSON (remove `"` or `'` if it has JSON in first and last each JSONs)
    json_detail = json.loads(detail_row)
    json_detail_data = json_detail["data"]
    # Count Array
    data_len = len(json_detail_data)

    if data_len != 0:
        for item in range(0, data_len):
            # print(json_detail_data[item])
            # standard_parameter col
            standard_parameter = list(json_detail_data[item].keys())[0]
            std_para.append(standard_parameter)
            # value col
            value = list(json_detail_data[item].values())[0]['v']
            value_arr.append(value)
            # inlimit col
            inlimit = list(json_detail_data[item].values())[0]['inlimit']
            inlimit_arr.append(inlimit)

            # Another col
            id_arr.append(id[row])
            date.append(str(day[row]).split(" ")[0] + " " + str(time[row]))
            mau_bantudong_id_arr.append(mau_bantudong_id[row])
            purpose_id_arr.append(purpose_id[row])
            tnn_station_id_arr.append(tnn_station_id[row])
            qualityindex_id_arr.append(qualityindex_id[row])
    else:
        # std_para, value and inlimit col set empty
        std_para.append("")
        value_arr.append("")
        inlimit_arr.append("")

        # Another col
        id_arr.append(id[row])
        date.append(str(day[row]).split(" ")[0] + " " + str(time[row]))
        mau_bantudong_id_arr.append(mau_bantudong_id[row])
        purpose_id_arr.append(purpose_id[row])
        tnn_station_id_arr.append(tnn_station_id[row])
        qualityindex_id_arr.append(qualityindex_id[row])

# Change name file output
write_xlsx = pd.ExcelWriter("data/result.xlsx", engine="xlsxwriter")
table_row = pd.DataFrame({"id": id_arr,
                         "tnn_station_id": tnn_station_id_arr,
                         "qualityindex_id": qualityindex_id_arr,
                         "mau_bantudong_id": mau_bantudong_id_arr,
                         "purpose_id": purpose_id_arr,
                         "standard_parameter_id": std_para,
                         "date": date,
                         "value": value_arr,
                         "inlimit": inlimit_arr})
table_row.to_excel(write_xlsx, sheet_name='data', index=False)
write_xlsx.save()
