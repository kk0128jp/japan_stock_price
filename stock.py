import tkinter as tk
import yfinance as yf
import os
import openpyxl
import pandas as pd

# メイン処理
def main():
    # 証券コード取得
    ticker_symbol = code_textbox.get()
    # コードに.T(東証)付ける
    t_code = addT(ticker_symbol)
    # コードから会社名取得
    com_name = getComInfo(t_code)['longName']
    # 最新日の株価データ取得
    ## 取得するカラム
    col_list = ["Open", "Close", "High", "Low"]
    ## 1日のデータ取得
    day_data = getDaydata(t_code)[col_list]
    # インデックス整形
    new_index_data = resetIndex(day_data,day_data.index)
    # 前日比カラム追加
    add_comparison_col = addComparisonCol(new_index_data)
    # excelへ保存
    writeDataToCsv(ticker_symbol, com_name, add_comparison_col)
    # 前日比算出
    ## ファイルパス
    desktop_path = os.path.expanduser('~\\Desktop')
    folder_path = desktop_path + "\\株価\\{}".format(com_name)
    file_path = folder_path + '\\data.xlsx'
    get_file_data = valComparison(file_path)
    # 前日比未入力セルの削除
    delNotEnteredCompareCol(file_path)
    # 前日比入力済みデータを再入力
    saveEnteredCompareCol(file_path, get_file_data)
    
# コードに東証.Tをつける
def addT(code):
    return code + '.T'
    
# 会社データを取得
def getComInfo(code):
    try:
        com_info = yf.Ticker(code)
        return com_info.info
    except Exception as e:
        showInfo(e)

# 最新日の株価取得
def getDaydata(code):
    com = yf.Ticker(code)
    day_data = com.history(period="1d")
    return day_data

# インデックスの整形
def resetIndex(data_frame, index):
    old_index = pd.to_datetime(index)
    new_index = old_index.strftime('%m/%d')
    reset_index_data = data_frame.reset_index()
    drop_data_col = reset_index_data.drop(columns='Date')
    drop_data_col['new_date'] = new_index
    set_new_index = drop_data_col.set_index("new_date")
    return set_new_index

# エクセルにデータ書き込み
def writeDataToCsv(code, com_name, csv_data):
    # フォルダ作成
    ## デスクトップのパスを取得
    desktop_path = os.path.expanduser('~\\Desktop')
    ## デスクトップ配下のフォルダパス
    folder_path = desktop_path + "\\株価\\{}".format(com_name)
    file_path = folder_path + '\\data.xlsx'
    if os.path.exists(folder_path) == False:
        # フォルダが存在しなかったら作成 
        os.makedirs(folder_path)
    # 事前にファイル,フォーマット作成
    if os.path.exists(file_path) == False:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws["A1"] = "社名"
        ws["B1"] = com_name
        ws["A2"] = "証券コード"
        ws["B2"] = code
        ws["A3"] = "取得日"
        ws["B3"] = "始値"
        ws["C3"] = "前日比"
        ws["D3"] = "終値"
        ws["E3"] = "前日比"
        ws["F3"] = "高値"
        ws["G3"] = "前日比"
        ws["H3"] = "安値"
        ws["I3"] = "前日比"
        wb.save(file_path)
    # xlsxファイルにデータ追記
    try:
        with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay') as writer:
            # 最終行取得
            load_file = openpyxl.load_workbook(file_path)
            sheet = load_file['Sheet']
            max_row = sheet.max_row
            # 最終行に追記
            csv_data.to_excel(writer, sheet_name='Sheet', startrow=max_row, header=False)
        showInfo("データを保存しました")
    except Exception as e:
        showInfo(e)
        
# 前日比カラムの追加
def addComparisonCol(data):
    # 追加するカラム
    cols = ["Open", "Close", "High", "Low"]
    # カラムの位置
    col_num = 1
    for i in cols:
        data.insert(col_num, i + '前日比', "→")
        col_num += 2
    return data

# 前日比を割り出す
def valComparison(file_path):
    # 保存されているデータ全て取り出す
    df = pd.read_excel(file_path, sheet_name="Sheet", header=None, skiprows=3)
    # データ行数取得
    data_row = len(df)
    cols = [1, 3, 5, 7]
    start_comp_col = 2
    # 前日比算出
    for i in cols:
        for j in range(data_row):
            #　基準日と次の日の値取得
            ## 最終行
            if (j == data_row - 1):
                # 次の日
                next_day_val = df.iat[j, i]
                # 前日比入力
                if (day_val > next_day_val):
                    df.iat[j, start_comp_col] = "↓"
                elif (day_val < next_day_val):
                    df.iat[j, start_comp_col] = "↑"
            else:
                # 基準日
                day_val = df.iat[j, i]
                # 次の日
                next_day_val = df.iat[j + 1, i]
                # 前日比入力
                if (day_val > next_day_val):
                    # >の場合
                    df.iat[j + 1, start_comp_col] = "↓"
                elif (day_val < next_day_val):
                    # <の場合
                    df.iat[j + 1, start_comp_col] = "↑"
        start_comp_col += 2
    return df

# 前日比未入力データを削除
def delNotEnteredCompareCol(file_path):
    wb = openpyxl.load_workbook(file_path)
    ws = wb['Sheet']
    
    # データの最終行を取得
    max_data_row = ws.max_row
    # データ行を削除
    ws.delete_rows(4, max_data_row)
    # 保存
    wb.save(file_path)

# 前日比カラムが入力済みのデータを保存
def saveEnteredCompareCol(file_path, data):
    with pd.ExcelWriter(file_path, mode="a", if_sheet_exists="overlay") as writer:
        data.to_excel(writer, sheet_name="Sheet", index=False, header=False, startrow=3)

# インフォメーションウィンドウ表示
def showInfo(error_messages):
    # サブウィンドウ
    sub_window = tk.Toplevel()
    sub_window.geometry('300x300')
    # エラーメッセージ表示
    sub_error_message = tk.Label(sub_window, text=error_messages)
    sub_error_message.place(x=30, y=70)
    # サブウィンドウ閉じる
    def closeWindow():
        sub_window.destroy()
    # 閉じるボタン表示
    sub_button = tk.Button(master=sub_window, text="閉じる", command=closeWindow)
    sub_button.place(x=30, y=100)

# 証券コードを保存するサブウィンドウを作成
def createTsSaveWindow():
    sub_window = tk.Toplevel()
    sub_window.geometry('500x500')
    label = tk.Label(sub_window, text="証券コード")
    label.place(x=30, y=70)

# メインウィンドウを作成
baseGround = tk.Tk()
# ウィンドウのサイズを設定
baseGround.geometry('600x600')
# 画面タイトル
baseGround.title('株式データ取得')

# ラベル
code_label = tk.Label(text='証券コード')
code_label.place(x=30, y=70)

# テキストボックス
code_textbox = tk.Entry(width=40)
code_textbox.place(x=30, y=90)

# 取得ボタン
button = tk.Button(baseGround, text='取得', command=main).place(x=30, y=120)

# 証券コードの保存の表示するボタン
code_setting_button = tk.Button(baseGround, text='証券コードの設定', command=createTsSaveWindow).place(x=30, y=170)

baseGround.mainloop()