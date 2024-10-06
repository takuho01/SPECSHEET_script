import pandas as pd
import numpy as np

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

# 使用例
file_path = 'CODE/SPECSHEET_script/sample.xlsx'
sheet_names = ['Sheet1', 'Sheet2']  # 複数のシート名をリストで指定
data_table_list = []

for sheet_index, sheet_name in enumerate(sheet_names):
    data_frame = read_excel_sheet(file_path, sheet_name)

    if data_frame is not None:
        print(f'sheet {sheet_index} data_frame : \n{data_frame}')

        # find key position
        for i in range(len(data_frame)):
            if data_frame.iloc[i, 0] == 'key':
                key_position = [i, 0]
                print(f'key position : {key_position}')
                break

        base_start_pos = [0, 0]
        for i in range(len(data_frame)):
            if data_frame.iloc[i, key_position[1]] == 'start':
                base_start_pos[0] = i - 1
                break
        for i in range(len(data_frame)):
            if data_frame.iloc[key_position[0], i] == 'start':
                base_start_pos[1] = i - 1
                break
        print(f'base_start_pos : {base_start_pos}')

        # find n2n, n2n_to pos
        for i in range(len(data_frame)):
            if data_frame.iloc[i, 0] == 'n2n':
                n2n_pos = [i, 0]
                print(f'n2n position : {n2n_pos}')
                break
        for i in range(len(data_frame)):
            if data_frame.iloc[i, 0] == 'n2n_to':
                n2n_to_pos = [i, 0]
                print(f'n2n_to position : {n2n_to_pos}')
                break

        # Create the data_table using base_start_pos as the header row and column
        header_row = base_start_pos[0]
        header_col = base_start_pos[1]

        # Extract the headers for columns
        column_headers = data_frame.iloc[header_row, header_col + 1:].values

        # Extract the headers for rows
        row_headers = data_frame.iloc[header_row + 1:, header_col].values

        # Extract the data for the table
        data = data_frame.iloc[header_row + 1:, header_col + 1:].values
        # add n2n, n2n_to data
        n2n_data = data_frame.iloc[n2n_pos[0], header_col + 1:].values
        n2n_to_data = data_frame.iloc[n2n_to_pos[0], header_col + 1:].values
        data = np.vstack([n2n_data, n2n_to_data, data])

        row_headers = np.hstack([['n2n'], ['n2n_to'], row_headers])

        # Create a DataFrame for the data_table
        data_table = pd.DataFrame(data, index=row_headers, columns=column_headers)
        data_table_list.append(data_table)

        print("Data Table:")
        print(data_table)

class DataEntry:
    def __init__(self, init_name, tgt_y_list, n2n_y_list, done_flag=False):
        self.init_name = init_name
        self.tgt_y_list = tgt_y_list
        self.n2n_y_list = n2n_y_list
        self.done_flag = done_flag

    def __repr__(self):
        return f"DataEntry(init_name={self.init_name}, tgt_y_list={self.tgt_y_list}, n2n_y_list={self.n2n_y_list}, done_flag={self.done_flag})"

init_list = []

for data_table in data_table_list:
    for index, row in data_table.iterrows():
        # if index is not n2n or n2n_to
        if index not in ['n2n', 'n2n_to']:
            init_name = index
            # add column data which is y value
            tgt_y_list = []
            for col in data_table.columns:
                if data_table.loc[index, col] == 'y':
                    tgt_y_list.append(col)
            # add n2n_to column data which is n2n y value
            n2n_y_list = []
            for col in data_table.columns:
                if data_table.loc['n2n', col] == 'y':
                    n2n_y_list.append(data_table.loc['n2n_to', col])
            done_flag = False
            entry = DataEntry(init_name, tgt_y_list, n2n_y_list, done_flag)
            init_list.append(entry)

print("Init List:")
for entry in init_list:
    print(entry)

# test 

