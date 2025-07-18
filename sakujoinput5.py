import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Tkinterの初期化
root = tk.Tk()
root.withdraw()  # メインウィンドウを非表示にする

# 入力用のExcelファイル(video_analyze2.pyで出力したもの)を選択
input_file_path = filedialog.askopenfilename(title="入力用のExcelファイルを選択", filetypes=[("Excel files", "*.xlsx;*.xls")])
if not input_file_path:
    print("入力用のファイルが選択されませんでした。")
    exit()

# 削除したい単語を指定
#target_word = input("削除したい単語を入力してください: ")
target_word = "その他"

# 入力ファイルを読み込む
df = pd.read_excel(input_file_path)

# 指定した単語と完全一致するセルが含まれる行を削除
df_cleaned = df[~df.apply(lambda row: row.astype(str).eq(target_word).any(), axis=1)]

# 結果を新しいExcelファイルに保存
output_file_path = filedialog.asksaveasfilename(title="結果を保存するファイル名を指定", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
if output_file_path:
    df_cleaned.to_excel(output_file_path, index=False)
    print(f"結果を {output_file_path} に保存しました。")
else:
    print("保存先のファイルが指定されませんでした。")

# 辞書用のExcelファイルを選択(今回は制御基板のファイル)
dict_file_path = filedialog.askopenfilename(title="辞書用のExcelファイルを選択", filetypes=[("Excel files", "*.xlsx;*.xls")])
if not dict_file_path:
    print("辞書用のファイルが選択されませんでした。")
    exit()

# 辞書を読み込む
dict_df = pd.read_excel(dict_file_path)
# 辞書の列名を指定
word_column_dict = '工程'  # 辞書の単語列名
mean_column_dict = '平均'  # 辞書の平均値列名
std_column_dict = '標準偏差'  # 辞書の標準偏差列名

# 辞書を辞書型に変換
word_stats = dict(zip(dict_df[word_column_dict], zip(dict_df[mean_column_dict], dict_df[std_column_dict])))

# 入力用のExcelファイルを選択
input_file_path = filedialog.askopenfilename(title="入力用のExcelファイルを選択", filetypes=[("Excel files", "*.xlsx;*.xls")])
if not input_file_path:
    print("入力用のファイルが選択されませんでした。")
    exit()

# 入力ファイルを読み込む
input_df = pd.read_excel(input_file_path)
# 単語が記入されている列の名前を指定（
word_column_input = 'timelinelabels'  # 入力ファイルの単語列名
measured_column_input = 'time'

# 結果を格納するための新しいDataFrameを作成
results = pd.DataFrame()
results[word_column_input] = input_df[word_column_input]
results[measured_column_input] = input_df[measured_column_input]


# 各単語に対する平均値、平均値＋標準偏差、平均値−標準偏差を計算
results['平均値'] = results[word_column_input].map(lambda x: word_stats[x][0] if x in word_stats else None)
results['平均値＋標準偏差'] = results[word_column_input].map(lambda x: word_stats[x][0] + word_stats[x][1] if x in word_stats else None)
results['平均値−標準偏差'] = results[word_column_input].map(lambda x: word_stats[x][0] - word_stats[x][1] if x in word_stats else None)
results['実測値'] = results[measured_column_input]

# 累積値を計算するための新しい列を追加
results['累積平均値'] = results['平均値'].cumsum()
results['累積平均値＋標準偏差'] = results['平均値＋標準偏差'].cumsum()
results['累積平均値−標準偏差'] = results['平均値−標準偏差'].cumsum()
results['累積実測値'] = results['実測値'].cumsum()

# 結果を新しいExcelファイルに保存
output_file_path = filedialog.asksaveasfilename(title="結果を保存するファイル名を指定", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx;*.xls")])
if output_file_path:
    results.to_excel(output_file_path, index=False)
    print(f"結果を {output_file_path} に保存しました。")
else:
    print("保存先のファイルが指定されませんでした。")
