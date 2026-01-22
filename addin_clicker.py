import win32com.client
import sys

def click_excel_addin_button(button_caption):
    try:
        # 起動中のExcelに接続
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("エラー: Excelが起動していません。")
        print("必ずExcelを起動し、対象のファイルを開いた状態でこのツールを実行してください。")
        # 処理終了（inputで止めるかどうかはバッチ側でも制御できるが、ここでもreturn）
        return

    import os
    import re

    cwd = os.getcwd()
    # カレントディレクトリの特定のパターンに一致するファイルを探す
    pattern = re.compile(r"^\d{5}_D-3\.xlsm$")
    import time
    
    target_files = [f for f in os.listdir(cwd) if pattern.match(f)]
    
    # Helper to get data from A.xlsx
    def get_data_from_a_file(app, a_file_path):
        print(f"参照ファイルを確認中: {a_file_path}")
        if not os.path.exists(a_file_path):
            print("参照ファイルが見つかりません。デフォルト回数(1回)で実行します。")
            return []
        
        try:
            wb_a = app.Workbooks.Open(a_file_path, ReadOnly=True)
            try:
                # シート "様式A-2" を検索
                try:
                    ws = wb_a.Sheets("様式A-2")
                except:
                    print(f"シート '様式A-2' が {a_file_path} に見つかりません。")
                    return []
                
                data_list = []
                row = 9 # B9 starting row
                while True:
                    # Column 2 is B, Column 3 is C
                    val_b = ws.Cells(row, 2).Value
                    
                    # Check for empty based on B column
                    if val_b is None or str(val_b).strip() == "":
                        break
                        
                    val_c = ws.Cells(row, 3).Value # C column
                    data_list.append((val_b, val_c))
                    row += 1
                
                print(f"データ行数をカウントしました: {len(data_list)}行")
                return data_list
            finally:
                wb_a.Close(SaveChanges=False)
        except Exception as e:
            print(f"参照ファイルの読み込み中にエラー: {e}")
            return []

    # Main file processing loop
    for filename in target_files:
        # Match again to be safe/clean
        m = pattern.match(filename)
        if not m: continue
        
        # Build A file path
        obj_id = filename[:5]
        a_filename = f"{obj_id}_A.xlsx"
        a_fullpath = os.path.join(cwd, a_filename)
        
        # 1. Get data
        data_rows = get_data_from_a_file(excel, a_fullpath)
        
        # Adjust loop count: total_rows (user request reverted the -1 logic)
        # Data writing loop should still range over len(data_rows)
        
        click_count = len(data_rows)
        # if click_count < 0: click_count = 0 # Safety not really needed if len >=0 but valid
        
        print(f"データ行数: {len(data_rows)}, 追加実行回数: {click_count}")

        if len(data_rows) == 0:
            print("追加すべき行数が0のため、スキップします。")
            continue

        # Check if D-3 file is actively open in Excel
        # Simple name check
        target_wb = None
        try:
            for wb in excel.Workbooks:
                if wb.Name == filename:
                    target_wb = wb
                    break
        except:
            pass

        if not target_wb:
            print(f"警告: 対象ファイル '{filename}' が開かれていません。")
            print("アドインを正しく動作させるため、先にExcelでファイルを開いてから実行してください。")
            continue
        
        print(f"対象ファイルが開かれていることを確認しました: {filename}")
        target_wb.Activate()
        
        # Find and click button 'click_count' times
        button_found_once = False
        target_control = None
        
        print(f"ボタン '{button_caption}' を探索中...")
        for bar in excel.CommandBars:
            if not bar.Visible and not bar.Enabled: continue
            try:
                for control in bar.Controls:
                    if control.Caption == button_caption:
                        target_control = control
                        break
                    if control.Type in [10, 14]:
                        try:
                            for sub in control.Controls:
                                if sub.Caption == button_caption:
                                    target_control = sub
                                    break
                        except: pass
                    if target_control: break
            except: pass
            if target_control: break
        
        if target_control:
            print(f"ボタンが見つかりました。{click_count}回 実行します。")
            
            for i in range(click_count):
                print(f"シート作成実行中... ({i+1}/{click_count})")
                try:
                    target_control.Execute()
                    time.sleep(2) 
                except Exception as e:
                    print(f"実行中にエラーが発生しました: {e}")
                    break
            
            # データの転記処理
            print("データの転記を開始します...")
            for i, (val_b, val_c) in enumerate(data_rows):
                # シート名: 様式D-3-{i+1}
                sheet_name = f"様式D-3-{i+1}"
                try:
                    ws_target = target_wb.Sheets(sheet_name)
                    print(f"転記中: {sheet_name} (M7={val_b}, Q7={val_c})")
                    ws_target.Range("M7").Value = val_b
                    ws_target.Range("Q7").Value = val_c
                except Exception as e:
                    print(f"シート '{sheet_name}' への転記失敗: {e}")

            button_found_once = True
        else:
            print(f"エラー: ボタン '{button_caption}' が見つかりませんでした。")

    # End of file loop
    print("全ファイルの処理が完了しました。")
    return


if __name__ == "__main__":
    # 画像にあるボタン名を指定
    target_button = "様式D-3シート追加" 
    click_excel_addin_button(target_button)
