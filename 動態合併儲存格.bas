Attribute VB_Name = "�ʺA�X���x�s��"
Option Explicit

Sub �ʺA�X��()
Application.DisplayAlerts = False '�@�~�t�δ�����r�A�Y�S�]�w�|�̭ȴ���
Dim i, j As Long '�ŧii�̫�Aj�H�`��ơAi���̫�@�C�Aj����e�C����
Dim myrng As Range '�ŧi�d���ܼ�
i = Cells(Rows.Count, 1).End(xlUp).Row '�ʺA�M��A��즳��Ƴ̫�@�C���C����
MsgBox "A��즳��Ƴ̫�@�C����" & i
For j = i To 2 Step -1 '�q�̫�@�C��ĤG�C����Astep -1 ���˼�
    Set myrng = Cells(j, "A") '�ثe�d��
    If myrng = myrng.Offset(-1, 0) Then '�Y�ثe��A���ȩM�e�@�C�ۦP
        myrng.Offset(-1, 0).Resize(2, 1).Merge '�h�ݥѤU�ӤW�X��
    End If
Next
Application.DisplayAlerts = True '���s�}�Ҧ۰ʴ�����r
End Sub
