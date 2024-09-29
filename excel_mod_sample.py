
import pandas as pd

# Excelファイルを読み込む関数
def read_excel_sheet(file_path, sheet_name):
    try:
        # Excelファイルを読み込む
        # header=Noneにしないと、一番上の行をデータとして読み取ってくれないので注意。
        df = pd.read_excel(file_path, sheet_name=sheet_name,header=None)
        return df
    except Exception as e:
        print(f"Error reading the Excel file: {e}")
        return None

# 使用例
file_path = 'CODE/SPECSHEET_script/sample.xlsx'
sheet_name = 'Sheet1'
data_frame = read_excel_sheet(file_path, sheet_name)

if data_frame is not None:
    print(data_frame)

print(data_frame.iloc[0,0])
if data_frame is not None:
    print(f'column 0 : \n{data_frame.iloc[:,0]}')
if data_frame is not None:
    print(f'row 0 : \n{data_frame.iloc[0, :]}')




