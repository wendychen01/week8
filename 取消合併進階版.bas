Attribute VB_Name = "取消合併進階版"
Option Explicit


Sub 取消合併2()
Dim shtsIdx As Integer
For shtsIdx = 1 To Sheets.Count '從第一
Sheets(shtsIdx).Activate '啟動

'剛剛取消合併單張表的Code
Dim rowCnt, mergeRow As Long
Dim myrng As Range '宣告範圍變數
rowCnt = Sheets(shtsIdx).UsedRange.Rows.Count 'rowCnt=列數

For Each myrng In Range(Cells(2, "A"), Cells(rowCnt, "A")) '從A2到A欄最後一列，逐列執行
    myrng.Select '選取範圍
    mergeRow = myrng.MergeArea.Count '合併範圍列數
    'MsgBox "目前是" & mergeRow & "列合併"
    myrng.UnMerge '取消合併
    myrng.Resize(mergeRow, 1) = myrng '*****補回原值
Next
Sheets(shtsIdx).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous  '給予框線
'End of 剛剛取消合併單張表的Code
Next
End Sub

