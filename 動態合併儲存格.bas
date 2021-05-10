Attribute VB_Name = "動態合併儲存格"
Option Explicit

Sub 動態合併()
Application.DisplayAlerts = False '作業系統提醒文字，若沒設定會依值提醒
Dim i, j As Long '宣告i最後，j違常整數，i為最後一列，j為當前列索引
Dim myrng As Range '宣告範圍變數
i = Cells(Rows.Count, 1).End(xlUp).Row '動態尋找A欄位有資料最後一列的列索引
MsgBox "A欄位有資料最後一列索引" & i
For j = i To 2 Step -1 '從最後一列到第二列遞減，step -1 為倒數
    Set myrng = Cells(j, "A") '目前範圍
    If myrng = myrng.Offset(-1, 0) Then '若目前的A欄位值和前一列相同
        myrng.Offset(-1, 0).Resize(2, 1).Merge '則需由下而上合併
    End If
Next
Application.DisplayAlerts = True '重新開啟自動提醒文字
End Sub
