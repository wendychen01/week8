Attribute VB_Name = "產能標記"
Option Explicit

Sub 產能標記()
Dim i, rowCnt As Integer
Dim tagetValue As Integer
tagetValue = CInt(InputBox("請輸入標記上限值(0-1000)"))
Dim rangeStr As String

rowCnt = Cells(Rows.Count, 1).End(xlUp).Row '最後一列
rangeStr = "b3:b" & rowCnt
MsgBox "目前運算範圍" & rangeStr
Range(rangeStr).Interior.Color = xlNone '運算範圍先還原為無顏色

For i = 3 To rowCnt '從第三列到最後一列
    If Cells(i, "B") > tagetValue Then
        Cells(i, "B").Interior.Color = vbYellow '背景給予黃色
    End If
Next
Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous

End Sub
