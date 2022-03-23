# coding: UTF-8
from email.policy import default
import os
from glob import glob
from tkinter import messagebox
import PySimpleGUI as sg

# os ファイル操作
# glob 指定フォルダ内のファイルを一括取得
# tkinter　メッセージボックス　https://pg-chain.com/python-messagebox
# PySimpleGUI　入力画面　　https://pysimplegui.readthedocs.io/en/latest/

# *******************************************
#  グローバル変数定義
# *******************************************
# -------------------------------------------
# モード
# -------------------------------------------
MODE_ADD = 0
MODE_REPLACE = 1
MODE_DEL = 2
# -------------------------------------------
# カラムの定義
# -------------------------------------------
col_add1 =[
  [sg.Text('挿入位置')],
  [sg.Text('挿入文字列')],
]
col_add2 =[
  [sg.InputText(key='-NUMBER1_ADD-', enable_events=True, size=(5, 1))],
  [sg.InputText(key='-STRING_ADD-',size=(20, 1))],
]
col_replace1 = [
  [sg.Text('開始位置')],
  [sg.Text('文字数')],
  [sg.Text('置換文字列')],  
]
col_replace2 = [
  [sg.InputText(key='-NUMBER1_REPLACE-', enable_events=True, size=(5, 1))],
  [sg.InputText(key='-NUMBER2_REPLACE-', enable_events=True, size=(5, 1))],
  [sg.InputText(key='-STRING_REPLACE-',size=(20, 1))],  
]
col_del1 = [
  [sg.Text('開始位置')],
  [sg.Text('文字数')],
]
col_del2 = [
  [sg.InputText(key='-NUMBER1_DEL-', enable_events=True, size=(5, 1))],
  [sg.InputText(key='-NUMBER2_DEL-', enable_events=True, size=(5, 1))],
]
# -------------------------------------------
# タブの定義
# -------------------------------------------
#　挿入
layout_AddTab = [
  [sg.Column(col_add1, justification='l'), sg.Column(col_add2, justification='l')]
]
#　編集
layout_ChangeTab = [
  [sg.Column(col_replace1, justification='l'), sg.Column(col_replace2, justification='l')]
]
#　削除
layout_DelTab = [
  [sg.Column(col_del1, justification='l'), sg.Column(col_del2, justification='l')]
]

# *******************************************
#  現在フォルダ内のファイル名を一括変更
# *******************************************  
def main():
  # -------------------------------------------
  # 入力画面を表示　＆　取得
  # -------------------------------------------
  # テーマカラー指定
  sg.theme('Dark Blue 3')

  # 全体レイアウト
  layout = [
    [sg.Text("フォルダ"), sg.InputText(key="FOLDER1"), sg.FolderBrowse('選択')],
    [sg.TabGroup([[
        sg.Tab('挿入', layout_AddTab, key="tab_add"),
        sg.Tab('編集', layout_ChangeTab, key="tab_replace"),
        sg.Tab('削除', layout_DelTab, key="tab_del")
    ]], key="tab_group", enable_events=True)],
    [sg.Button('実行', key='-ACT-'), sg.Exit('終了')]
  ]

  # ウィンドウ定義　画面サイズ変更可
  window = sg.Window('ファイル名変更ツール', layout, resizable=True)

  #GUI表示実行部分
  while True:
    # ウィンドウ
    event, values = window.read()
      # -------------------------------------------
      # 終了ボタン押下
      # -------------------------------------------   
    if event == sg.WIN_CLOSED or event == '終了':
      break

      # -------------------------------------------
      # タブ切り替え
      # -------------------------------------------   
    elif event=="tab_group":
        select_tab = values["tab_group"]
        img_name = ""

        if select_tab=="tab_add":
          # 別タブの入力値クリア
          window['-NUMBER1_REPLACE-'].update('')
          window['-NUMBER2_REPLACE-'].update('')
          window['-STRING_REPLACE-'].update('')
          window['-NUMBER1_DEL-'].update('')
          window['-NUMBER2_DEL-'].update('')
        elif select_tab=="tab_replace":
          # 別タブの入力値クリア
          window['-NUMBER1_ADD-'].update('')
          window['-STRING_ADD-'].update('')
          window['-NUMBER1_DEL-'].update('')
          window['-NUMBER2_DEL-'].update('')
        elif select_tab=="tab_del":
          # 別タブの入力値クリア
          window['-NUMBER1_ADD-'].update('')
          window['-STRING_ADD-'].update('')
          window['-NUMBER1_REPLACE-'].update('')
          window['-NUMBER2_REPLACE-'].update('')
          window['-STRING_REPLACE-'].update('')

      # -------------------------------------------
      # 実行ボタン押下
      # -------------------------------------------   
    elif event == '-ACT-':     
        # -------------------------------------------
        # 値取得＆チェック
        # -------------------------------------------             
        try:
          errtitle=''
          errmsg=''
          # 値取得
          select_tab = values["tab_group"]
          fldpath = values['FOLDER1']
          if select_tab == "tab_add":
            startnum=values['-NUMBER1_ADD-']
            str=values['-STRING_ADD-']
          elif select_tab == "tab_replace":
            startnum=values['-NUMBER1_REPLACE-']
            countstr=values['-NUMBER2_REPLACE-']
            str=values['-STRING_REPLACE-']            
          else:
            startnum=values['-NUMBER1_DEL-']
            countstr=values['-NUMBER2_DEL-']

          # 値チェック
          # フォルダパス
          errmsg = chkfldpath(fldpath)
          if errmsg != '':
            errtitle = 'フォルダパスエラー'     
            errmsg = 'フォルダパス\n' + errmsg              
            raise Exception('チェックエラー')    

          #### 挿入 ####
          if select_tab == "tab_add":
            # 数値チェック
            errmsg = chknum(startnum)
            if errmsg != '':
              errtitle = '開始位置エラー'
              errmsg = '挿入位置\n' + errmsg
              raise Exception('チェックエラー')   
            if int(startnum) < 0:
              errtitle = '挿入位置エラー'
              errmsg = '挿入位置\n' + '0以上で入力してください。'           
              raise Exception('チェックエラー')                    
            # 挿入文字列チェック
            errmsg = chkstr(str)
            if errmsg != '':
              errtitle = '置き換え文字列エラー'       
              errmsg = '挿入文字列:\n' + errmsg                                 
              raise Exception('チェックエラー')

          #### 編集 ####              
          elif select_tab == "tab_replace":
            # 編集開始位置チェック
            errmsg = chknum(startnum)
            if errmsg != '':
              errtitle = '開始位置エラー'
              errmsg = '開始位置\n' + errmsg              
              raise Exception('チェックエラー')    
            if int(startnum) < 1:
              errtitle = '開始位置エラー'
              errmsg = '開始位置\n' + '1以上で入力してください。'           
              raise Exception('チェックエラー')                  
            # 編集文字数チェック
            errmsg = chknum(countstr)
            if errmsg != '':
              errtitle = '文字数エラー'        
              errmsg = '文字数\n' + errmsg              
              raise Exception('チェックエラー')       
            if int(countstr) < 1:
              errtitle = '文字数エラー'
              errmsg = '文字数\n' + '1以上で入力してください。'           
              raise Exception('チェックエラー')                                
            # 置換文字チェック
            errmsg = chkstr(str)
            if errmsg != '':
              errtitle = '置き換え文字列エラー'       
              errmsg = '置換文字列\n' + errmsg              
              raise Exception('チェックエラー')  

          #### 削除 ####                    
          else:
            # 開始位置チェック
            errmsg = chknum(startnum)
            if errmsg != '':
              errtitle = '開始位置エラー'
              errmsg = '開始位置\n' + errmsg              
              raise Exception('チェックエラー')    
            if int(startnum) < 1:
              errtitle = '開始位置エラー'
              errmsg = '開始位置\n' + '1以上で入力してください。'           
              raise Exception('チェックエラー')               
            # 編集文字数チェック
            errmsg = chknum(countstr)
            if errmsg != '':
              errtitle = '文字数エラー'        
              errmsg = '文字数\n' + errmsg              
              raise Exception('チェックエラー')   
            if int(startnum) < 1:
              errtitle = '文字数エラー'
              errmsg = '文字数\n' + '1以上で入力してください。'           
              raise Exception('チェックエラー')                  

          # ファイルリスト取得
          path_list = glob(os.path.join(fldpath,'*.*'))

          # -------------------------------------------
          # 処理実行
          # -------------------------------------------     
          #### 挿入 ####          
          if select_tab == "tab_add":
              # ファイル名変更処理
              ret = changefilename_add(path_list,int(startnum),str)

          #### 編集 ####    
          elif select_tab == "tab_replace":
              # ファイル名変更処理
              ret = changefilename_replace(path_list,int(startnum), int(countstr),str)

          #### 削除 ####   
          else:
            # 削除
              ret = changefilename_del(path_list,int(startnum), int(countstr))
              
        #### チェックエラー　####
        except Exception as e:    
          print(e)
          if len(errmsg) == 0:
            errtitle='システムエラー'
            errmsg='システムエラー' + e

          messagebox.showerror(errtitle, errmsg)
  window.close()


# *******************************************
#  関数定義
# ******************************************* 
# -------------------------------------------
# フォルダパスチェック
# -------------------------------------------
def chkfldpath(path):
  if len(path) == 0:
    # 空白チェック
    msg = "フォルダを指定してください。"
    return msg
  if os.path.exists(path) == False:
    # 存在チェック
    msg = "指定フォルダが存在しません。"
    return msg
  
  return ''

# -------------------------------------------
# 数値チェック
# -------------------------------------------
def chknum(strnum):
  if not strnum:
    # 空白チェック
    msg = "値を入力してください。"
    return msg
  if not strnum.isnumeric():
    # 数値チェック
    msg = "数値を入力してください。"
    return msg
  return ''
# 文字数チェック

# -------------------------------------------
# 文字列チェック
# -------------------------------------------
def chkstr(str):
  if not str:
    # 空白チェック
    msg = "値を入力してください。"
    return msg
  return ''

# -------------------------------------------
# ファイル名変更_挿入
# -------------------------------------------
def changefilename_add(path_list,startnum,stradd):
# 指定フォルダ内のファイルパスリストを取得
  for path in path_list:
      # フォルダとファイル名取得
      dirname, file_ext = os.path.split(path)
      # ファイル名と拡張子取得
      file, ext = os.path.splitext(file_ext)

      # 前方文字列と後方文字列と追加文字列から新しいファイル名を作成
      strbefore = file[0:startnum]
      strafter = file[startnum:] 
      new_file_name = strbefore + stradd + strafter + ext
      # ファイル名書き換え
      os.rename(path, os.path.join(dirname, new_file_name))


# -------------------------------------------
# ファイル名変更_編集
# -------------------------------------------
def changefilename_replace(path_list,startnum, countstr,strreplace):
# 指定フォルダ内のファイルパスリストを取得
  for path in path_list:
      # フォルダとファイル名取得
      dirname, file_ext = os.path.split(path)
      # ファイル名と拡張子取得
      file, ext = os.path.splitext(file_ext)

      # 前方文字列と後方文字列と置き換え文字列から新しいファイル名を作成
      strbefore = file[0:startnum - 1]
      strafter = file[startnum + countstr - 1:] 
      new_file_name = strbefore + strreplace + strafter + ext
      # ファイル名書き換え
      os.rename(path, os.path.join(dirname, new_file_name))

# -------------------------------------------
# ファイル名変更_削除
# -------------------------------------------
def changefilename_del(path_list,startnum, countstr):
# 指定フォルダ内のファイルパスリストを取得
  alldeletecount = 0
  for path in path_list:
      # フォルダとファイル名取得
      dirname, file_ext = os.path.split(path)
      # ファイル名と拡張子取得
      file, ext = os.path.splitext(file_ext)

      # 前方文字列と後方文字列から削除文字を除いて新しいファイル名を作成
      strbefore = file[0:startnum - 1]
      strafter = file[startnum + countstr - 1:] 
      if len(strbefore) == 0 and len(strafter) == 0:
        # 指定文字数がファイル名を超えていた場合は、'-'とする
        alldeletecount = alldeletecount + 1
        new_file_name = '-' + str(alldeletecount) + ext
      else:
        new_file_name = strbefore  + strafter + ext
      # ファイル名書き換え
      os.rename(path, os.path.join(dirname, new_file_name))



# *******************************************
#  メイン処理実行
# ******************************************* 
if __name__ == "__main__":
    main()

