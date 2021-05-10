Attribute VB_Name = "計算列數"
Option Explicit

Sub 計算列數()
Dim rowCnt As Integer
rowCnt = Sheets(1).UsedRange.Rows.Count
MsgBox "目前已使用列數" & rowCnt
End Sub
