import pandas as pd
import sys
import io
import json
import chardet
from datetime import datetime, timedelta

# Set up UTF-8 encoding for stdout and stderr
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# Error handling function to return JSON
def return_error_json(error_message):
    result = {
        'error': True,
        'message': error_message
    }
    print(json.dumps(result))
    sys.exit(1)

# Function to detect file encoding
def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        result = chardet.detect(file.read(10000))  # Analyze a portion of the file
        return result['encoding']

# Get file paths from command-line arguments
input_file = sys.argv[1]
standard_times_file = sys.argv[2]

# Detect the encoding of the input file
try:
    input_file_encoding = detect_encoding(input_file)
except Exception as e:
    return_error_json(f"Error detecting encoding of input file: {e}")

# Read the CSV files
try:
    score_df = pd.read_csv(input_file, encoding=input_file_encoding)
    standard_times_df = pd.read_csv(standard_times_file, encoding='utf-8')  # Always UTF-8 for standard times
except FileNotFoundError as e:
    return_error_json(f"File not found: {e}")
except pd.errors.ParserError as e:
    return_error_json(f"Error parsing CSV: {e}")
except Exception as e:
    return_error_json(f"An unexpected error occurred: {e}")

# 列名とデータから空白や特殊文字を削除
score_df.columns = score_df.columns.str.strip()
standard_times_df.columns = standard_times_df.columns.str.strip()

# `apply`と`map`を使用してデータから空白や特殊文字を削除
score_df = score_df.apply(lambda x: x.map(lambda y: y.strip() if isinstance(y, str) else y))
standard_times_df = standard_times_df.apply(lambda x: x.map(lambda y: y.strip() if isinstance(y, str) else y))

# hh:mm:ss形式の時間を秒数に変換する関数
def time_to_seconds(time_str):
    if pd.isna(time_str) or time_str == '':
        return 0
    h, m, s = map(int, time_str.split(':'))
    return h * 3600 + m * 60 + s

# 秒数をhh:mm:ss形式に変換する関数
def seconds_to_time(seconds):
    if pd.isna(seconds):
        return '00:00:00'
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02}:{m:02}:{s:02}"

# 学習開始日時に標準学習時間を足す関数
def add_standard_time(start_time_str, standard_time_seconds):
    # 時刻データを変換、エラー時はNaTを返す
    start_time = pd.to_datetime(start_time_str, errors='coerce')
    
    if pd.isna(start_time) or pd.isna(standard_time_seconds):
        return ''
    
    end_time = start_time + timedelta(seconds=int(standard_time_seconds))
    return end_time.strftime('%Y/%m/%d %H:%M:%S')

# 所要時間と標準学習時間を秒数に変換
score_df['所要時間（秒）'] = score_df['所要時間'].apply(time_to_seconds)
standard_times_df['標準学習時間（秒）'] = standard_times_df['標準学習時間'].apply(time_to_seconds)

# 総標準学習時間を計算
total_standard_time_seconds = standard_times_df['標準学習時間（秒）'].sum()

# マージ
merged_df = pd.merge(score_df, standard_times_df, on='コンテンツ名', how='left')
merged_df['フォルダ名'] = score_df['フォルダ名']  # フォルダ名列を保持

# CSV読み込み後の処理で日時の変換を行う部分
merged_df['学習開始日時'] = pd.to_datetime(merged_df['学習開始日時'], errors='coerce')

# コンテンツ名と氏名でグループ化して、合計所要時間を計算
grouped = merged_df.groupby(['氏名', 'コンテンツ名']).agg({
    '所要時間（秒）': 'sum',
    '標準学習時間（秒）': 'first'
}).reset_index()

# 合計結果が数値型の0の場合にその行を削除
def filter_invalid_rows(grouped_df):
    # 合計所要時間が数値型の0でないことを条件に設定
    conditions = (grouped_df['所要時間（秒）'] != 0)
    # 条件を満たす行のみを返す
    return grouped_df[conditions]

# 確認時間とマークを設定する関数
def set_mark_and_confirm_time(row, merged_df):
    total_time = row['所要時間（秒）']
    standard_time = row['標準学習時間（秒）']
    mark = 'O' if total_time >= standard_time else 'X'
    sub_df = merged_df[(merged_df['氏名'] == row['氏名']) & (merged_df['コンテンツ名'] == row['コンテンツ名'])]
    sub_df = sub_df.sort_values(by='学習開始日時')

    # sub_dfが空である場合のチェック
    if sub_df.empty:
        print(f"Warning: No matching data for {row['氏名']} and {row['コンテンツ名']}")
        return merged_df

    if '理解度テスト' in row['コンテンツ名']:
        confirm_time = min(total_time, standard_time)
        merged_df.loc[sub_df.index, '確認時間'] = seconds_to_time(confirm_time)
        merged_df.loc[sub_df.index, 'マーク'] = 'O'
    else:
        if mark == 'O':
            merged_df.loc[sub_df.index, '確認時間'] = sub_df['標準学習時間（秒）'].apply(seconds_to_time)
        else:
            merged_df.loc[sub_df.index, '確認時間'] = '00:00:00'
        merged_df.loc[sub_df.index, 'マーク'] = mark

    if mark == 'X':
        total_time_sum = sub_df['所要時間（秒）'].sum()
        missing_time = max(0, standard_time - total_time_sum)  # 足りない時間を計算

        # 足りない時間が発生する場合にhh:mm:ss形式に変換して設定
        if missing_time > 0:
            merged_df.loc[sub_df.index, '足りない時間'] = seconds_to_time(missing_time)
        else:
            merged_df.loc[sub_df.index, '足りない時間'] = '00:00:00'

        # 残りの行の足りない時間は空欄にする
        merged_df.loc[sub_df.index[1:], '足りない時間'] = ''

    return merged_df

# 各グループに対してマークと確認時間を設定
for idx, row in grouped.iterrows():
    merged_df = set_mark_and_confirm_time(row, merged_df)

# 'マーク'列が作成されているか確認
if 'マーク' not in merged_df.columns:
    merged_df['マーク'] = ''  # 初期化

# 足りない時間を計算
merged_df['足りない時間（秒）'] = merged_df['標準学習時間（秒）'] - merged_df['所要時間（秒）']
merged_df['足りない時間'] = merged_df.apply(lambda row: seconds_to_time(row['足りない時間（秒）']) if row['マーク'] == 'X' and row['足りない時間（秒）'] > 0 else '', axis=1)

# 学習完了日を計算
merged_df['学習完了日'] = merged_df.apply(lambda row: add_standard_time(row['学習開始日時'], row['標準学習時間（秒）']) if row['マーク'] == 'O' else '', axis=1)

# 必要な列を選択
columns_to_output = [
    'グループ', 'グループ(全階層)', '氏名', 'フォルダ名', 'コンテンツ名',
    '学習開始日時', '学習完了日', '所要時間', '標準学習時間', '確認時間', 
    'マーク', '足りない時間', 'URL'
]

# 存在しない列を確認し、警告を出す
existing_columns = [col for col in columns_to_output if col in merged_df.columns]
missing_columns = [col for col in columns_to_output if col not in merged_df.columns]
if missing_columns:
    print(f"Warning: The following columns are missing in the merged data: {missing_columns}")

# ここでoutput_dfを初期化
output_df = merged_df.loc[:, existing_columns]

# standard_timesのコンテンツ名の順序に基づいて並び替え
# Get unique content names from standard_times_df, remove nulls and check if it's empty
content_order = standard_times_df['コンテンツ名'].drop_duplicates().dropna().tolist()

# Ensure content_order is not empty
if not content_order:
    print("Warning: No valid categories found in 'コンテンツ名'")
    content_order = output_df['コンテンツ名'].drop_duplicates().tolist()  # Use output_df's content if needed

# Only create Categorical if content_order is valid
if content_order:
    output_df['コンテンツ順序'] = pd.Categorical(output_df['コンテンツ名'], categories=content_order, ordered=True)
else:
    output_df['コンテンツ順序'] = output_df['コンテンツ名']  # Fallback to using raw content

# 学習開始日時を datetime 型に変換 - Allow flexible date formats
output_df['学習開始日時'] = pd.to_datetime(output_df['学習開始日時'], errors='coerce')

# Optional: Log any parsing errors (NaT values) to a file
with open('debug_log.txt', 'w') as f:
    f.write(f"NaT entries in '学習開始日時': {merged_df['学習開始日時'].isna().sum()}\n")
    f.write(merged_df[merged_df['学習開始日時'].isna()].to_string())

# 氏名でグループ化し、コンテンツ順序と学習開始日時で並び替え
output_df = output_df.sort_values(['氏名', 'コンテンツ順序', '学習開始日時']).drop(columns=['コンテンツ順序'])

# 学習開始日時を文字列型に戻す (use the original format 'YYYY/MM/DD HH:MM')
output_df['学習開始日時'] = output_df['学習開始日時'].dt.strftime('%Y/%m/%d %H:%M')

# DataFrame のデータ型を JSON シリアライズ可能な型に変換
output_df = output_df.astype(str)

# 結果をJSONとして出力
result = {
    'data': output_df.to_dict(orient='records'),
    'totalStandardTime': int(total_standard_time_seconds)
}
print(json.dumps(result))


