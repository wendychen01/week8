Attribute VB_Name = "����аO"
Option Explicit

Sub ����аO()
Dim i, rowCnt As Integer
Dim tagetValue As Integer
tagetValue = CInt(InputBox("�п�J�аO�W����(0-1000)"))
Dim rangeStr As String

rowCnt = Cells(Rows.Count, 1).End(xlUp).Row '�̫�@�C
rangeStr = "b3:b" & rowCnt
MsgBox "�ثe�B��d��" & rangeStr
Range(rangeStr).Interior.Color = xlNone '�B��d����٭쬰�L�C��

For i = 3 To rowCnt '�q�ĤT�C��̫�@�C
    If Cells(i, "B") > tagetValue Then
        Cells(i, "B").Interior.Color = vbYellow '�I����������
    End If
Next
Range("a1").CurrentRegion.Borders.LineStyle = xlContinuous

End Sub
