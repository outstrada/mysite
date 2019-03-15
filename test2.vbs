' Excel起動
Set oXlsApp = CreateObject("Excel.Application")
If oXlsApp Is Nothing Then
 ' Excel起動失敗
 MsgBox "Excel起動失敗"
Else
 ' Excel起動成功
 ' --Excel表示（falseにすると非表示にできる）
 oXlsApp.Application.Visible = true
 ' --Excelの警告を非表示にする
 oXlsApp.Application.DisplayAlerts = False
 ' --3秒待つ
 ' WScript.Sleep(3000)
 ' --ブック追加
 ' oXlsApp.Application.Workbooks.Add()
 ' ブックを開く
 oXlsApp.Application.Workbooks.Open("C:\Users\TOMOHIRO\Desktop\プログラミング\vb\Test.xlsm")
 ' --シート選択
 Set oSheet = oXlsApp.Worksheets(1)
 ' --A1のセルに値を設定
 ' oSheet.Range("A1").value = "aaa"
 ' ' --行が2、列が3のセルに値を設定
 ' oSheet.Cells(2, 3).value = 100
' セルをアクティブ
 oSheet.Range("D150").Activate
 ' --3秒待つ
 ' WScript.Sleep(3000)
 ' --Excel終了
 ' oXlsApp.Quit
 ' --Excelオブジェクトクリア
 Set oXlsApp = Nothing
End If
