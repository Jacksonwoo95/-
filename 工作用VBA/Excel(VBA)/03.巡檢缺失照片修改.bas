Attribute VB_Name = "���˯ʥ�"
Sub �D�{��()
Attribute �D�{��.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' ����Ӥ��վ�
    Call �Ӥ��վ�
    
    ' ��������V�P�_�M���
    Call �P�_�h�ӹϧΪ����V
    
    ' ����ĤG���Ӥ��վ�
    ' Call �Ӥ��վ�_2
End Sub

Sub �Ӥ��վ�()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A3:E23")
    
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
        ActiveSheet.Shapes.SelectAll
        
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
        ActiveSheet.Shapes.SelectAll
        Selection.ShapeRange.Select False
        
        ' ���s��ܤ��O�¤W�δ¤U������
        If �O�d��ܪ�����.Count > 0 Then
            �O�d��ܪ�����(1).Select
            For i = 2 To �O�d��ܪ�����.Count
                �O�d��ܪ�����(i).Select False
            Next i
            ' MsgBox "�w��� " & �O�d��ܪ�����.Count & " �Ӥ��¤W�δ¤U������C"
        Else
            ' MsgBox "�S����줣�¤W�δ¤U������C"
        End If
    Else
        ' MsgBox "�u�@���S������C"
    End If
End Sub

Sub �Ӥ��վ�_2()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A3:E23")
    
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

Sub ��X�ʥ���r()
    Dim result As String
    Dim lastRow As Long
    Dim i As Long
    Dim lineResult As String
    
    ' �M�� M1 �x�s��
    Range("M1").Value = ""
    
    ' ��X�̫�@�ӻݭn�ˬd����
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' �ˬd�ós�����w�x�s�檺��
    For i = 1 To lastRow Step 23  ' �C 23 ���ˬd�@��
        lineResult = ConcatIfNotEmpty(Range("A" & i)) & _
                     ConcatIfNotEmpty(Range("F" & i)) & _
                     ConcatIfNotEmpty(Range("A" & (i + 1))) & _
                     ConcatIfNotEmpty(Range("F" & (i + 1)))
        
        ' �p�G�o�@�զ�����D�ŭȡA�K�[�쵲�G���ô���
        If Len(lineResult) > 0 Then
            result = result & Left(lineResult, Len(lineResult) - 1) & vbNewLine
        End If
    Next i
    
    ' �������G�r�꥽��������š]�p�G�����ܡ^
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If
    
    ' �N���G�g�J M1 �x�s��
    Range("M1").Value = result
    
    ' �]�m�x�s��榡���۰ʴ���
    Range("M1").WrapText = True
End Sub

Function ConcatIfNotEmpty(cell As Range) As String
    If Not IsEmpty(cell) Then
        ConcatIfNotEmpty = cell.Value & " "
    Else
        ConcatIfNotEmpty = ""
    End If
End Function


