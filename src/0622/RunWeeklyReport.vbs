Dim xlApp, xlBook
Set xlApp = CreateObject("Excel.Application")

' 裏でひっそり実行させる設定
xlApp.Visible = False
xlApp.DisplayAlerts = False

' 【修正済み】ユーザー様の環境のパスと実際のファイル名を設定しました
' ※この行が途中で改行されないよう、1行で記述してください
Set xlBook = xlApp.Workbooks.Open("C:\Users\busia\ドキュメント\AQC\src\0630\Send_Mail_Studio.xlsm")

' Excel内のメール送信マクロを実行
xlApp.Run "定期メール送信実行"

' 保存してExcelを完全に閉じる
xlBook.Close True
xlApp.Quit

Set xlBook = Nothing
Set xlApp = Nothing