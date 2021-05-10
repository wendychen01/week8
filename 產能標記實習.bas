Attribute VB_Name = "產能標記實習"
Option Explicit

Sub 產能標記實習()
Dim i, rowCnt As Integer
Dim tagetValueUB, tagetValueLB As Integer '實習功能-上下變數
tagetValueUB = CInt(InputBox("請輸入標記上限值(0-1000)"))
tagetValueLB = CInt(InputBox("請輸入標記下限值(0-1000)"))
Dim rangeStr As String

rowCnt = Cells(Rows.Count, 1).End(xlUp).Row '最後一列
rangeStr = "b3:b" & rowCnt
MsgBox "目前運算範圍" & rangeStr
Range(rangeStr).Interior.Color = xlNone '運算範圍先還原為無顏色

For i = 3 To rowCnt '從第三列到最後一列
    If Cells(i, "B") > tagetValueUB Then
        Cells(i, "B").Interior.Color = vbYellow '上限背景給予黃色
    End If
    If Cells(i, "B") < tagetValueLB Then
        Cells(i, "B").Interior.Color = vbBlue '下限背景給予藍色
    End If
    
Next
Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous

End Sub
