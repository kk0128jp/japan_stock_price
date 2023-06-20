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
    ## 始値
    open_price = getDaydata(t_code)["Open"]
    ## 終値
    close_price = getDaydata(t_code)["Close"]
    ## 高値
    high_price = getDaydata(t_code)["High"]
    ## 安値
    low_price = getDaydata(t_code)["Low"]
    ## まとめたデータ
    day_data = getDaydata(t_code)[["Open","Close","High","Low"]]
    # インデックス整形
    old_index = pd.to_datetime(close_price.index)
    new_index = old_index.strftime('%m/%d')
    reset_index_data = close_price.reset_index()
    drop_data_col = reset_index_data.drop(columns='Date')
    drop_data_col['new_date'] = new_index
    set_new_index = drop_data_col.set_index("new_date")
    # csv保存
    writeDataToCsv(ticker_symbol, com_name, set_new_index)
    
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

#def writeDataToCsv(csv_data):
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
        ws["A1"] = code
        ws["B1"] = com_name
        ws["A2"] = "取得日"
        ws["B2"] = "始値"
        ws["C2"] = "終値"
        ws["D2"] = "高値"
        ws["E2"] = "安値"
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

# インフォメーションウィンドウ表示
def showInfo(error_messages):
    # サブウィンドウ
    sub_window = tk.Toplevel()
    sub_window.geometry('300x300')
    # エラーメッセージ表示
    sub_error_message = tk.Label(sub_window, text=error_messages)
    sub_error_message.place(x=30, y=70)
    # 閉じるボタン表示
    sub_button = tk.Button(master=sub_window, text="閉じる")
    sub_button.place(x=30, y=100)
    
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

# 取得するカラムの指定
check_label = tk.Label(text='保存する項目')
check_label.place(x=30, y=120)
## デフォルトで全てオン
col_def_value = tk.BooleanVar()
col_def_value.set(True)
# チェックボックス表示
## 始値 
open_checkbox = tk.Checkbutton(baseGround, variable=col_def_value, text='始値')
open_checkbox.place(x=30, y=140)
## 終値
close_checkbox = tk.Checkbutton(baseGround, variable=col_def_value, text='終値')
close_checkbox.place(x=90, y=140)
## 高値
high_checkbox = tk.Checkbutton(baseGround, variable=col_def_value, text='高値')
high_checkbox.place(x=150, y=140)
## 安値
low_checkbox = tk.Checkbutton(baseGround, variable=col_def_value, text='安値')
low_checkbox.place(x=210, y=140)

# 取得ボタン
button = tk.Button(baseGround, text='取得', command=main).place(x=30, y=180)

baseGround.mainloop()