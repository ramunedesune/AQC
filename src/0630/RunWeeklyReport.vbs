Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")

xlApp.Visible = False
xlApp.DisplayAlerts = False

' 実際のファイルパスに書き換えてください（1行で記述します）
Set xlBook = xlApp.Workbooks.Open("C:\Users\busia\ドキュメント\AQC\src\0630\master.xlsm")

' 上記で作成したマクロを実行
xlApp.Run "定期メール送信実行"

xlBook.Close True
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing