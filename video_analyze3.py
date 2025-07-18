import json
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib import font_manager
from tkinter import Tk, StringVar, Radiobutton, Button, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
import japanize_matplotlib  # 日本語フォントの設定を簡単にするライブラリ
import numpy as np  # 分散計算のためにNumPyをインポート
import os  # ファイルパス操作のためにosモジュールをインポート

# GUIでファイルパスを指定
def select_file():
    Tk().withdraw()  # Tkinterのウィンドウを非表示にする
    file_path = askopenfilename(title="JSONファイルを選択", filetypes=[("JSON files", "*.json")])
    return file_path

# JSONファイルを読み込む
def load_json(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return json.load(file)

# videoを選択する
def select_video(data):
    videos = [item['video'] for item in data if 'video' in item]
    unique_videos = list(set(videos))  # 重複を排除
    if not unique_videos:
        messagebox.showinfo("情報", "ビデオが見つかりません。")
        return None

    # Tkinterウィンドウを作成
    root = Tk()
    root.title("ビデオ選択")

    # 選択されたビデオを保持する変数
    selected_video = StringVar(root)

    # ラジオボタンを作成
    for video in unique_videos:
        Radiobutton(root, text=video, variable=selected_video, value=video).pack(anchor='w')

    # 確定ボタン
    def confirm_selection():
        if selected_video.get():
            root.quit()  # ウィンドウを閉じる
        else:
            messagebox.showerror("エラー", "ビデオが選択されていません。")

    confirm_button = Button(root, text="選択", command=confirm_selection)
    confirm_button.pack(pady=20)

    root.mainloop()  # GUIのメインループを開始

    return selected_video.get()

# データを抽出して昇順にソート
def extract_and_sort_data(data, selected_video):
    extracted_data = []
    for item in data:
        if item.get("video") == selected_video and "videoLabels" in item and item["videoLabels"]:
            for label in item["videoLabels"]:
                if "ranges" in label and label["ranges"]:
                    for range_item in label["ranges"]:
                        start = range_item.get("start")
                        end = range_item.get("end")
                        timelinelabels = label.get("timelinelabels", [])
                        
                        # timelinelabelsをカンマ区切りの文字列に変換
                        timelinelabels_str = ', '.join(timelinelabels)

                        # 抽出したデータをリストに追加
                        extracted_data.append({
                            "start": start,
                            "end": end,
                            "timelinelabels": timelinelabels_str  # 文字列に変換して保存
                        })
    
    # startで昇順に並び替え
    extracted_data.sort(key=lambda x: x['start'])
    # 追加の計算
    if extracted_data:
        first_start = extracted_data[0]['start']
        fps = 30  # フレームレート

        for item in extracted_data:
            start_0 = item['start'] - first_start
            end_0 = item['end'] - first_start
            timelinelabels_0 = item['timelinelabels']
            timelinelabels_1 = timelinelabels_0  # timelinelabelsのコピー
            timelinelabels_2 = timelinelabels_0  # timelinelabelsのコピー
            time_start = start_0 / fps
            time_end = end_0 / fps
            time = time_end - time_start

            item.update({
                "start_0": start_0,
                "end_0": end_0,
                "timelinelabels_0": timelinelabels_0,
                "timelinelabels_1": timelinelabels_1,
                "timelinelabels_2": timelinelabels_2,
                "time_start": time_start,
                "time_end": time_end,
                "time": time
            })

        # sum_timeをtime_endの最大値に変更
        max_time_end = max(item['time_end'] for item in extracted_data)
        for item in extracted_data:
            item['sum_time'] = max_time_end

        # timelinelabelsごとの合計時間、平均、分散、標準偏差を計算
        label_time_summary = {}
        for item in extracted_data:
            label = item['timelinelabels']
            if label not in label_time_summary:
                label_time_summary[label] = {'times': [], 'total_time': 0, 'count': 0}
            label_time_summary[label]['times'].append(item['time'])
            label_time_summary[label]['total_time'] += item['time']
            label_time_summary[label]['count'] += 1

        # 合計時間、平均、分散、標準偏差を追加
        for label, summary in label_time_summary.items():
            total_time = summary['total_time']
            average_time = total_time / summary['count']
            variance = np.var(summary['times'])  # 分散を計算
            std_dev = np.std(summary['times'])  # 標準偏差を計算
            summary['average_time'] = average_time
            summary['variance'] = variance
            summary['std_dev'] = std_dev  # 標準偏差を追加

        # 各アイテムに合計時間、平均、分散、標準偏差を追加
        for item in extracted_data:
            label = item['timelinelabels']
            item['label_total_time'] = label_time_summary[label]['total_time']
            item['label_average_time'] = label_time_summary[label]['average_time']
            item['label_variance'] = label_time_summary[label]['variance']
            item['label_std_dev'] = label_time_summary[label]['std_dev'] 
    return extracted_data

#データをExcelファイルに保存
def save_to_excel(data):
    # 名前をつけて保存ダイアログを表示
    save_path = asksaveasfilename(title="Excelファイルを保存", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        # 指定された順番で列を作成
        columns_order = [
            "start", "end", "timelinelabels",
            "start_0", "end_0", "timelinelabels_0",
            "time_start", "time_end", "timelinelabels_1",
            "time", "timelinelabels_2", "sum_time",
            "label_total_time", "label_average_time", "label_variance", "label_std_dev"
        ]
        
        # DataFrameを作成し、指定された順番で列を並べ替え
        df = pd.DataFrame(data)[columns_order]
        df.to_excel(save_path, index=False)
        print(f"データが {save_path} に保存されました。")
        
        # エクセル名を取得
        excel_name = os.path.splitext(os.path.basename(save_path))[0]
        return excel_name  # エクセル名を返す

    return None  # 保存しなかった場合


#平均値・標準偏差のみExcelファイルに保存
"""def save_to_excel2(data):
    # 名前をつけて保存ダイアログを表示
    save_path2 = asksaveasfilename(title="Excelファイルを保存", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if save_path2:
        #必要な列のみ抜き取り
        selected_columns = [
            "timelinelabels", "label_total_time", "label_average_time", "label_variance", "label_std_dev"
        ]

        # DataFrameを作成し、指定された順番で列を並べ替え
        df = pd.DataFrame(data)[selected_columns]
        df.to_excel(save_path2, index=False)
        print(f"データが {save_path2} に保存されました。")
        
        # エクセル名を取得
        excel_name = os.path.splitext(os.path.basename(save_path2))[0]
        return excel_name  # エクセル名を返す

    return None  # 保存しなかった場合"""



# グラフを描画
def plot_graphs(data, excel_name):
    # ガントチャートの描画
    plt.figure(figsize=(10, 6))
    
    for item in data:
        bar = plt.barh(item['timelinelabels'], item['time_end'] - item['time_start'], left=item['time_start'], edgecolor='black')
        
        # 各バーの中央にtimeの値を表示
        plt.text(item['time_start'] + (item['time_end'] - item['time_start']) / 2, 
                 item['timelinelabels'], 
                 f"{item['time']:.2f}", 
                 ha='center', va='center', color='white')

    plt.xlabel('時間 (秒)')
    plt.ylabel('ラベル')
    plt.title('ガントチャート')
    plt.grid(axis='x')
    plt.savefig(f"{excel_name}_1.png")  # ガントチャートを保存
    plt.show()

    # 円グラフの描画
    labels = [item['timelinelabels'] for item in data]
    total_times = [item['label_total_time'] for item in data]
    average_times = [item['label_average_time'] for item in data]
    variances = [item['label_variance'] for item in data]
    std_devs = [item['label_std_dev'] for item in data]

    # 重複ラベルを集約
    label_summary = {}
    for label, total_time, average_time, variance, std_dev in zip(labels, total_times, average_times, variances, std_devs):
        if label not in label_summary:
            label_summary[label] = {'total_time': total_time, 'average_time': average_time, 'variance': variance, 'std_dev': std_dev}
        else:
            # 既に存在する場合は、合計時間はそのまま、平均、分散、標準偏差は上書き
            label_summary[label]['average_time'] = average_time
            label_summary[label]['variance'] = variance
            label_summary[label]['std_dev'] = std_dev

    # 円グラフのデータ準備
    pie_labels = list(label_summary.keys())
    pie_sizes = [summary['total_time'] for summary in label_summary.values()]
    pie_explode = [0.1] * len(pie_labels)  # 各セクションを少しずらす

    # 円グラフの描画
    plt.figure(figsize=(8, 8))
    wedges, texts, autotexts = plt.pie(pie_sizes, explode=pie_explode, labels=pie_labels, autopct='', startangle=140)
    
    # ラベルに合計時間、平均時間、分散、標準偏差、割合を追加
    for i, (text, summary) in enumerate(zip(autotexts, label_summary.values())):
        percentage = (summary['total_time'] / sum(pie_sizes)) * 100  # 割合を計算
        text.set_text(f"{pie_labels[i]}\n合計: {summary['total_time']:.2f}秒\n平均: {summary['average_time']:.2f}秒\n分散: {summary['variance']:.2f}\n標準偏差: {summary['std_dev']:.2f}\n割合: {percentage:.1f}%")
        text.set_fontsize(8)  # フォントサイズを小さくする

    plt.title('ラベルごとの合計時間、平均時間、分散、標準偏差、割合')
    plt.axis('equal')  # 円を正しく表示
    plt.savefig(f"{excel_name}_2.png")  # 円グラフを保存
    plt.show()

def plot_mean_with_error_bars(data, excel_name):
    import numpy as np
    import matplotlib.pyplot as plt

    labels = []
    means = []
    std_devs = []

    # タイムラインラベルごとの平均と標準偏差を集める
    label_summary = {}
    for item in data:
        label = item['timelinelabels']
        if label not in label_summary:
            label_summary[label] = {
                'total_time': 0,
                'count': 0,
                'times': []
            }
        label_summary[label]['total_time'] += item['time']
        label_summary[label]['count'] += 1
        label_summary[label]['times'].append(item['time'])

    for label, summary in label_summary.items():
        labels.append(label)
        means.append(summary['total_time'] / summary['count'])
        std_devs.append(np.std(summary['times']))

    # グラフ作成
    x = np.arange(len(labels))
    width = 0.4

    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.bar(x, means, width, yerr=std_devs, capsize=5, color='lightblue', label='平均 ± 標準偏差')

    # 軸ラベルとタイトル
    ax.set_xlabel('タイムラインラベル')
    ax.set_ylabel('時間 (秒)')
    ax.set_title('タイムラインラベルごとの平均と標準偏差')

    # x軸のラベルを斜め45度に回転させて見やすく
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)  # フォントサイズを小さく

    ax.legend()
    ax.yaxis.grid(True)

    # 平均値と標準偏差の数値を表示
    for i, bar in enumerate(bars):
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval, f'{means[i]:.2f}', ha='center', va='bottom', fontsize=9, color='black')
        ax.text(bar.get_x() + bar.get_width()/2, yval - std_devs[i], f'±{std_devs[i]:.2f}', ha='center', va='top', fontsize=9, color='black')

    plt.tight_layout()  # レイアウト調整
    plt.savefig(f"{excel_name}_mean_with_error_bars.png")
    plt.show()


# メイン処理
def main():
    file_path = select_file()
    if file_path:
        data = load_json(file_path)
        if isinstance(data, list) and len(data) > 0:
            selected_video = select_video(data)
            if selected_video is not None:
                sorted_data = extract_and_sort_data(data, selected_video)                
                excel_name = save_to_excel(sorted_data)
                #excel_name2 = save_to_excel2(sorted_data)
                if excel_name:  # エクセル名が取得できた場合
                    plot_graphs(sorted_data, excel_name)  # グラフを描画
                    plot_mean_with_error_bars(sorted_data, excel_name)  # エラーバー付きのグラフを描画
                
        else:
            print("データが空のリストです。")
    else:
        print("ファイルが選択されませんでした。")

if __name__ == "__main__":
    main()