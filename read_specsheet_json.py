import pandas as pd
import numpy as np
import json
# Excelファイルを読み込む関数
def read_excel_sheet(file_path, sheet_name):
    try:
        # Excelファイルを読み込む
        # header=Noneにしないと、一番上の行をデータとして読み取ってくれないので注意。
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        return df
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None


def read_excel_to_json(file_path, n2n_data):
    excel_sheet_data = {}
    sheet_names = ['Connections']

    data_frame = read_excel_sheet(file_path, sheet_names[0])
    sheet_data = process_connections_sheet(data_frame, n2n_data)
    excel_sheet_data['Connections_sheet'] = sheet_data

    return excel_sheet_data

def process_n2n_sheet(data_frame):
    sheet_data = {}
    for i in range(len(data_frame)):
        if data_frame.iloc[i,0] == 'key':
            key_position = [i, 0]
    for i in range(len(data_frame.columns)):
        if data_frame.iloc[key_position[0],i] == 'NOC':
            NOC_position = [key_position[0], i]
        if data_frame.iloc[key_position[0],i] == 'n2n_target':
            n2n_target_position = [key_position[0], i]
        if data_frame.iloc[key_position[0],i] == 'n2n_to':
            n2n_to_position = [key_position[0], i]

    n2n_list = []
    for i in range(key_position[0] + 1, len(data_frame)):
        n2n_list.append({
            'NOC': data_frame.iloc[i, NOC_position[1]],
            'n2n_target': data_frame.iloc[i, n2n_target_position[1]],
            'n2n_to': data_frame.iloc[i, n2n_to_position[1]]
        })

    sheet_data['n2n_data'] = n2n_list
    return sheet_data

def process_connections_sheet(data_frame, n2n_data):
    sheet_data = {}
    for i in range(len(data_frame)):
        if data_frame.iloc[i,0] == 'key':
            key_position = [i, 0]
        if data_frame.iloc[i,0] == 'target':
            target_position = [i, 0]
        if data_frame.iloc[i,0] == 'start':
            row_start_position = [i, 0]
    for i in range(len(data_frame.columns)):
        if data_frame.iloc[key_position[0],i] == 'initiator':
            initiator_position = [key_position[0], i]
        if data_frame.iloc[key_position[0],i] == 'start':
            col_start_position = [key_position[0], i]

    sheet_data['initiator_data'] = []
    for init in range(row_start_position[0], len(data_frame)):
        tgt_list = []
        tgt_yn_list = []
        n2n_list = []
        for tgt in range(col_start_position[1], len(data_frame.columns)):
            tgt_yn_list.append(data_frame.iloc[init, tgt])
            if data_frame.iloc[init, tgt] in ['y', 'Y']:
                tgt_list.append(data_frame.iloc[target_position[0], tgt])
            for n2n in n2n_data:
                if data_frame.iloc[init, tgt] in ['y', 'Y']:
                    if data_frame.iloc[target_position[0], tgt] == n2n['n2n_target']:
                        n2n_list.append(n2n['n2n_to'])
        if len(n2n_list) == 0: n2n_check_done_flag = True
        else: n2n_check_done_flag = False
        
        sheet_data['initiator_data'].append({
            'initiator_name': data_frame.iloc[init, initiator_position[1]],
            'target_list': tgt_list,
            'tgt_yn_list': [],
            'n2n_list': n2n_list,
            'connect_check_done_flag': n2n_check_done_flag
        })

    return sheet_data

def show_excel_data(excel_data):
    print("Excel Sheet Data:--------------------------")
    for key, value in excel_data.items():
        print(f"\n{key}:")
        if isinstance(value, dict):
            for sub_key, sub_value in value.items():
                print(f"  {sub_key}:")
                if isinstance(sub_value, list):
                    for item in sub_value:
                        print(f"    - {item}")
                else:
                    print(f"    {sub_value}")
        else:
            print(f"  {value}")

def process_memmap_sheet(data_frame):
    sheet_data = {}
    for i in range(len(data_frame)):
        if data_frame.iloc[i,0] == 'key':
            key_position = [i, 0]
    for i in range(len(data_frame.columns)):
        if data_frame.iloc[key_position[0],i] == 'NOC':
            NOC_position = [key_position[0], i]
        if data_frame.iloc[key_position[0],i] == 'target':
            target_position = [key_position[0], i]

    memmap_list = []
    for i in range(key_position[0] + 1, len(data_frame)):
        memmap_list.append({
            'NOC': data_frame.iloc[i, NOC_position[1]],
            'target': data_frame.iloc[i, target_position[1]],
        })

    sheet_data['memmap_data'] = memmap_list
    return sheet_data

def connect_check(initiator_data):
    initiator_data_mod = []
    for initiator in initiator_data:
        if initiator['connect_check_done_flag'] == False:
            for n2n in initiator['n2n_list']:
                for initiator02 in initiator_data:
                    if n2n == initiator02['initiator_name'] and initiator02['connect_check_done_flag'] == True:
                        initiator['target_list'].extend(initiator02['target_list'])
                        initiator['n2n_list'].remove(n2n)
            if len(initiator['n2n_list']) == 0:
                initiator['connect_check_done_flag'] = True
    initiator_data_mod.append(initiator)
    return initiator_data_mod

def gen_yn_matrix(initiator_data, memmap_data):
    for initiator in initiator_data:
        for memmap in memmap_data:
            yn_flag = 'n'
            for target in initiator['target_list']:
                if memmap['target'] == target:
                    yn_flag = 'y'
            initiator['tgt_yn_list'].append(yn_flag)
    return initiator_data

def write_excel_sheet(file_path, sheet_name, sheet_data):
    df = pd.DataFrame(sheet_data)
    df.to_excel(file_path,sheet_name=sheet_name, header=None,index=None)


top_sheet_data = {}

output_excel = read_excel_sheet('CODE/SPECSHEET_script/sample_top.xlsx','Memmap')
top_sheet_data['Memmap_sheet'] = process_memmap_sheet(output_excel)
output_excel = read_excel_sheet('CODE/SPECSHEET_script/sample_top.xlsx','n2n')
top_sheet_data['n2n_sheet'] = process_n2n_sheet(output_excel)

all_excel_data = {}

file_path_list = {
    'A': 'CODE/SPECSHEET_script/sampleA.xlsx', 
    'B': 'CODE/SPECSHEET_script/sampleB.xlsx',
    }


for key, value in file_path_list.items():
    excel_data = read_excel_to_json(value, top_sheet_data['n2n_sheet']['n2n_data'])
    all_excel_data[key] = excel_data
    show_excel_data(all_excel_data[key])




top_sheet_data['Connections_sheet'] = {'initiator_data': []}
for key, value in file_path_list.items():
    top_sheet_data['Connections_sheet']['initiator_data'].extend(all_excel_data[key]['Connections_sheet']['initiator_data'])

show_excel_data(top_sheet_data)

connect_check(top_sheet_data['Connections_sheet']['initiator_data'])
gen_yn_matrix(top_sheet_data['Connections_sheet']['initiator_data'], top_sheet_data['Memmap_sheet']['memmap_data'])

# tgt_yn_listをoutput_excelに出力
connect_output_data = [
    [memmap['NOC'] for memmap in top_sheet_data['Memmap_sheet']['memmap_data']],
    [memmap['target'] for memmap in top_sheet_data['Memmap_sheet']['memmap_data']]
]
for tgt_yn in top_sheet_data['Connections_sheet']['initiator_data']:
    connect_output_data.append(tgt_yn['tgt_yn_list'])

# Convert connect_output_data to a DataFrame
initiator_names = [initiator['initiator_name'] for initiator in top_sheet_data['Connections_sheet']['initiator_data']]
initiator_names.insert(0, 'NOC')
initiator_names.insert(1, 'Target')

# Create DataFrame from connect_output_data
df = pd.DataFrame(connect_output_data)
df_initiator_names = pd.DataFrame([initiator_names]).transpose()

# Combine df_initiator_names and df
df = pd.concat([df_initiator_names, df], axis=1)

# Update connect_output_data with the new DataFrame
connect_output_data = df.values.tolist()


write_excel_sheet('CODE/SPECSHEET_script/output_excel.xlsx', 'Connection_result', connect_output_data)

show_excel_data(top_sheet_data)



