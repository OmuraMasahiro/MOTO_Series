import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
from tkinter import messagebox
import web_sc
def folder_set():
    filetype = [("エクセル","*.xlsx")]
    file_path = filedialog.askopenfilename(filetypes = filetype)
    entry_import.insert(0, file_path)

def execute():
    path = entry_import.get()
    print(path)
    if(path==''):
        messagebox.showinfo("結果", "ファイルが選択されていません")
    else:
        web_sc.con_chrome(path)
        messagebox.showinfo("結果", "作成しました")

# rootメインウィンドウの設定
root = tk.Tk()
root.title("帳票自動作成 MOTO 1号")
root.geometry("600x100")

# メインフレームの作成と設置
frame = ttk.Frame(root)
frame.grid(row=0, column=0, sticky=tk.NSEW, padx=10, pady=10)

# 各種ウィジェットの作成
label_number = ttk.Label(frame, text="雇用保険被保険者資格喪失届を作成します")
label_export = ttk.Label(frame, text="入力ファイル(エクセル)：")
entry_import = ttk.Entry(frame, width=30)
button_export = ttk.Button(frame, text="検索", command=folder_set)

button_execute = ttk.Button(frame, text="実行", command=execute)

# 各種ウィジェットの設置
label_number.grid(row=0, column=0)
label_export.grid(row=1, column=0)
entry_import.grid(row=1, column=1)
button_export.grid(row=1, column=2, padx=5)

button_execute.grid(row=2, column=1, pady=3)


root.mainloop()