Attribute VB_Name = "�����X�ֶi����"
Option Explicit


Sub �����X��2()
Dim shtsIdx As Integer
For shtsIdx = 1 To Sheets.Count '�q�Ĥ@
Sheets(shtsIdx).Activate '�Ұ�

'�������X�ֳ�i��Code
Dim rowCnt, mergeRow As Long
Dim myrng As Range '�ŧi�d���ܼ�
rowCnt = Sheets(shtsIdx).UsedRange.Rows.Count 'rowCnt=�C��

For Each myrng In Range(Cells(2, "A"), Cells(rowCnt, "A")) '�qA2��A��̫�@�C�A�v�C����
    myrng.Select '����d��
    mergeRow = myrng.MergeArea.Count '�X�ֽd��C��
    'MsgBox "�ثe�O" & mergeRow & "�C�X��"
    myrng.UnMerge '�����X��
    myrng.Resize(mergeRow, 1) = myrng '*****�ɦ^���
Next
Sheets(shtsIdx).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous  '�����ؽu
'End of �������X�ֳ�i��Code
Next
End Sub

