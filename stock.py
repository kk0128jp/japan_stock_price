import tkinter as tk
import yfinance as yf

# メイン処理
def main():
    # コードに.T(東証)付ける
    t_code = addT(code_textbox.get())
    # コードからデータを取得
    print(getComInfo(t_code))

# コードに東証.Tをつける
def addT(code):
    return code + '.T'

# 株式データを取得
def getComInfo(code):
    try:
        com_info = yf.Ticker(code)
        raise ValueError("valueError")
        #return com_info.info
    except Exception as e:
        showErrorWindow(e)

# エラーウィンドウ表示
def showErrorWindow(error_messages):
    # サブウィンドウ
    sub_window = tk.Toplevel(master=None)
    sub_window.geometry('300x300')
    # エラーメッセージ表示
    sub_error_message = tk.Label(text=error_messages)
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