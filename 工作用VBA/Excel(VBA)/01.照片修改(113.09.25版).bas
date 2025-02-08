Attribute VB_Name = "�Ӥ��վ�v240925"
Sub �D�{��()
Attribute �D�{��.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' ����Ӥ��վ�
    Call �Ӥ��վ�
    
    ' ��������V�P�_�M���
    Call �P�_�h�ӹϧΪ����V
    
End Sub

Sub �Ӥ��վ�()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A1:E24")
    
    i = mergedRange.Height
    j = mergedRange.Width
    
    ws.Range("L1").Value = i
    ws.Range("M1").Value = j
    
    With Selection.ShapeRange
        .LockAspectRatio = False
        .Height = j
        .Width = i * 0.95
    End With
End Sub

Sub �P�_�h�ӹϧΪ����V()
    Dim shp As Shape
    Dim ���䨤�� As Double
    Dim ���ਤ�� As Double
    Dim ���G As String
    Dim �O�d��ܪ����� As New Collection
    
    ���G = ""
    
    If ActiveSheet.Shapes.Count > 0 Then
        ' ������ܩҦ�����
        
        For Each shp In Selection.ShapeRange
            ���ਤ�� = (shp.Rotation Mod 360 + 360) Mod 360
            
            If shp.Width >= shp.Height Then
                ���䨤�� = ���ਤ��
            Else
                ���䨤�� = (���ਤ�� + 90) Mod 360
            End If
            
            ���G = ���G & "�ϧ� " & shp.name & " �����䭱�V"
            
            If (���䨤�� >= 315 Or ���䨤�� <= 45) Then
                ���G = ���G & "�k�C" & vbNewLine
                �O�d��ܪ�����.Add shp
            ElseIf (���䨤�� > 45 And ���䨤�� <= 135) Then
                ���G = ���G & "�U�C" & vbNewLine
            ElseIf (���䨤�� > 135 And ���䨤�� <= 225) Then
                ���G = ���G & "���C" & vbNewLine
                �O�d��ܪ�����.Add shp
            ElseIf (���䨤�� > 225 And ���䨤�� < 315) Then
                ���G = ���G & "�W�C" & vbNewLine
            End If
        Next shp
        
        ' MsgBox ���G
        
        ' �����Ҧ����
        ' ActiveSheet.Shapes.SelectAll
        Selection.ShapeRange.Select False
        
        ' ���s��ܤ��O�¤W�δ¤U������
        If �O�d��ܪ�����.Count > 0 Then
            �O�d��ܪ�����(1).Select
            For i = 2 To �O�d��ܪ�����.Count
                �O�d��ܪ�����(i).Select False
            Next i
            Call �Ӥ��վ�_2
            ' MsgBox "�w��� " & �O�d��ܪ�����.Count & " �Ӥ��¤W�δ¤U������C"
        Else
            ' MsgBox "�S����줣�¤W�δ¤U������C"
        End If
    Else
        MsgBox "�u�@���S������C"
    End If
End Sub

Sub �Ӥ��վ�_2()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A1:E24")
    
    i = mergedRange.Height
    j = mergedRange.Width
    
    ws.Range("L1").Value = i
    ws.Range("M1").Value = j
    
    With Selection.ShapeRange
        .LockAspectRatio = False
        .Height = i * 0.95
        .Width = j
    End With
End Sub

Sub SetAlternatingBorderColors()
Attribute SetAlternatingBorderColors.VB_ProcData.VB_Invoke_Func = "e\n14"
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


Sub �R������ï�Ҧ��Ӥ�()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("�z�T�w�n�R���Ҧ��Ӥ���?", vbYesNo + vbQuestion, "�T�{�R��")
    
    If response = vbYes Then
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        
    End If
    ' �p�G��� "�_"�A�{���N���������A����ܥ���T��
    
End Sub
