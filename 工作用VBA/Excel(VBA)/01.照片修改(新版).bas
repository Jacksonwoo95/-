Attribute VB_Name = "���շӤ�"
Sub �Ӥ��վ�() '�Ӥ��̷�e��e�C���վ�
Attribute �Ӥ��վ�.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' �Ӥ��վ� ����
'
' �ֳt��: Ctrl+q
'
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range

    Set ws = ActiveSheet  ' �ϥη�e���ʪ��u�@��
    Set mergedRange = ws.Range("A1:E24")  ' �]�w���X�֪��d��

    ' ����X���x�s�檺�`���ס]���ס^�M�`�e��
    i = mergedRange.Height  ' �x�s�檺�`���ס]�Y���ס^
    j = mergedRange.Width  ' �x�s�檺�`�e��

    ' �N�ȿ�J L1 �M M1 �x�s��
    'ws.Range("L1").Value = i
    'ws.Range("M1").Value = j

    ' �վ��e����d�򪺤ؤo
    With Selection.ShapeRange
        .LockAspectRatio = False  ' ���\�e����ҧ���
        .Height = j '�e��
        .Width = i * 0.95   ' 0.95������
    End With
End Sub
Sub �Ӥ��վ�_��() '�Ӥ��̷�e��e�C���վ�(���e���)
Attribute �Ӥ��վ�_��.VB_ProcData.VB_Invoke_Func = "e\n14"

'
' �Ӥ��ϦV ����
'
' �ֳt��: Ctrl+e
'
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range

    Set ws = ActiveSheet  ' �ϥη�e���ʪ��u�@��
    Set mergedRange = ws.Range("A1:E24")  ' �]�w���X�֪��d��

    ' ����X���x�s�檺�`���ס]���ס^�M�`�e��
    i = mergedRange.Height  ' �x�s�檺�`���ס]�Y���ס^
    j = mergedRange.Width  ' �x�s�檺�`�e��

    ' �N�ȿ�J L1 �M M1 �x�s��
    'ws.Range("L1").Value = i
    'ws.Range("M1").Value = j

    ' �վ��e����d�򪺤ؤo
    With Selection.ShapeRange
        .LockAspectRatio = False  ' ���\�e����ҧ���
        .Height = i * 0.95    ' 0.95������
        .Width = j ' �e��
    End With
    
End Sub


Sub ����()
    Dim i As Integer
    For i = 1 To 3
        Selection.ShapeRange.IncrementRotation 90
    Next i
End Sub

Sub �R������ï�Ҧ��Ӥ�()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("�z�T�w�n�R���Ҧ��Ӥ���?", vbYesNo + vbQuestion, "�T�{�R��")
    
    If response = vbYes Then
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        
    End If
    ' �p�G��� "�_"�A�{���N���������A����ܥ���T��
    
End Sub

Sub SetAlternatingBorderColors()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim finalRow As Long
    Dim i As Integer

    ' �]�m��e�E�����u�@���ؼФu�@��
    Set ws = ActiveSheet

    ' ����ؼФu�@�����̫�@���
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    ' �p��̲צ�ơ]�b�̫�@�檺��¦�W�A�[�W52��^
    finalRow = lastRow + 208

    ' �H26�檺���j�B�z�Ҧ���������A����̲צ�
    For i = 25 To finalRow Step 26
        ApplyAlternatingBorders ws.Range("L" & i & ":O" & i)
        ApplyAlternatingBorders ws.Range("Q" & i & ":T" & i)
    Next i
End Sub

Sub ApplyAlternatingBorders(rng As Range)
    Dim cell As Range
    Dim borderColor As Long
    
    For Each cell In rng
        If (cell.Column - rng.Cells(1, 1).Column) Mod 2 = 0 Then
            borderColor = vbRed ' ���ƦC�]������
        Else
            borderColor = vbBlue ' �_�ƦC�]���Ŧ�
        End If
        
        ' �]�m�W�B�U�B���B�k���ؽu
        With cell.Borders
            .Item(xlEdgeTop).LineStyle = xlContinuous
            .Item(xlEdgeTop).Color = borderColor
            .Item(xlEdgeBottom).LineStyle = xlContinuous
            .Item(xlEdgeBottom).Color = borderColor
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeLeft).Color = borderColor
            .Item(xlEdgeRight).LineStyle = xlContinuous
            .Item(xlEdgeRight).Color = borderColor
        End With
    Next cell
End Sub


