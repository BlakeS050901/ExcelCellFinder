import os
import pandas as pd

counters = {
    "D_01": 0, "D_02A": 0, "D_06": 0, "D_11": 0, "D_16": 0, "D_18": 0,
    "D_26": 0, "D_27": 0, "D_04": 0, "D_05": 0, "D_05A": 0, "D_06A": 0,
    "D_07": 0, "D_07A": 0, "D_08": 0, "D_09": 0, "D_15": 0, "D_17": 0,
    "D_19": 0, "D_20": 0, "D_12": 0, "D_13": 0, "D_14": 0, "D_22": 0,
    "D_23": 0, "D_24": 0, "D_25": 0, "D_21": 0, "D_02": 0,
}

files = []
dir_string = 'Directory Here'
directory = os.fsencode(dir_string)

for file in os.listdir(directory):
    filename = os.fsdecode(file)
    if filename.endswith(".xlsx"):
        files.append(filename)
        continue
    else:
        continue

for file in files:
    file = dir_string + '/' + file
    df = pd.read_excel(file)
    for column in df.columns:
        for cell in df[column]:
            if pd.notnull(cell):
                data = str(cell)
                data = data.replace('-', '_')
                if data in counters:
                    counters[data] += 1


num = 0
for key in counters.keys():
    print(key, counters[key])
    num += counters[key]
print(num)


