import tkinter as tk
import yfinance as yf
import os

# メイン処理
def main():
    # コードに.T(東証)付ける
    t_code = addT(code_textbox.get())
    # 証券コード取得
    ticker_symbol = code_textbox.get()
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
    # csv保存
    writeDataToCsv(com_name,day_data)
    
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
def writeDataToCsv(com_name, csv_data):
    # フォルダ作成
    ## デスクトップのパスを取得
    desktop_path = os.path.expanduser('~\\Desktop')
    ## デスクトップ配下のフォルダパス
    folder_path = desktop_path + "\\株価\\{}".format(com_name)
    file_path = folder_path + '\\data.csv'
    if os.path.exists(folder_path) == False:
        # フォルダが存在しなかったら作成 
        os.makedirs(folder_path)
        #showInfo("デスクトップに「株価」フォルダを作成しました")
    # 事前にファイル,フォーマット作成
    if os.path.exists(file_path):
        with open(file_path, 'w') as file:
            file.write()
    # csvファイルにデータ追記
    try:
        csv_data.to_csv(file_path, mode='a', header=False)
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
    
# 取得ボタン
button = tk.Button(baseGround, text='取得', command=main).place(x=30, y=130)

baseGround.mainloop()