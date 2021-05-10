Attribute VB_Name = "計算欄數"
Option Explicit


Sub 計算欄數()
Dim colCnt As Integer
colCnt = Sheets(1).UsedRange.Columns.Count
MsgBox "目前已使用欄數" & colCnt
End Sub
