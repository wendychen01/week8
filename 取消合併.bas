Attribute VB_Name = "�����X��"
Option Explicit


Sub �����X��()
Dim rowCnt, mergeRow As Long
Dim myrng As Range '�ŧi�d���ܼ�
rowCnt = Sheets(1).UsedRange.Rows.Count 'rowCnt=�C��

For Each myrng In Range(Cells(2, "A"), Cells(rowCnt, "A")) '�qA2��A��̫�@�C�A�v�C����
    myrng.Select '����d��
    mergeRow = myrng.MergeArea.Count '�X�ֽd��C��
    MsgBox "�ثe�O" & mergeRow & "�C�X��"
    myrng.UnMerge '�����X��
    myrng.Resize(mergeRow, 1) = myrng '*****�ɦ^���
Next
Sheets(1).Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous  '�����ؽu

End Sub

