Attribute VB_Name = "���L�M��VBA1"
Sub �ƻs�Ĥ@���榡�s��3��()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' �]�w����e�u�@��

    Dim startRow As Long
    startRow = 1 ' �q��1��}�l

    Dim copyRange As Range
    Set copyRange = ws.Range("A1:I40")

    Dim loopCounter As Integer
    loopCounter = 0 ' �p�ƾ���l��

    ' �~��ƻs����F�� 10 ���ΨS����������ӽƻs���e
    Do While loopCounter < 3 And startRow <= ws.Rows.Count - 80 ' �T�O����������ӽƻs���e
        copyRange.Copy Destination:=ws.Range("A" & startRow + 40)
        startRow = startRow + 40
        loopCounter = loopCounter + 1
    Loop
End Sub

Sub �վ�Ӥ����e()
Attribute �վ�Ӥ����e.VB_ProcData.VB_Invoke_Func = "q\n14"

' �ֳt��: Ctrl+q
    Dim ws As Worksheet
    Set ws = ActiveSheet ' �]�w����e�u�@��
    
    
    Selection.Placement = xlMoveAndSize '�Ӥ���m�j�p�H�x�s�����
    
    ' ����X���x�s�檺�d��
    Dim mergedRange As Range
    Set mergedRange = ws.Range("B5:I20")

    ' �p��X���x�s�檺��e�`�M
    Dim i As Double
    Dim col As Range
    For Each col In mergedRange.Columns
        i = i + col.Width
    Next col

    ' �p��X���x�s�檺�C���`�M
    Dim j As Double
    Dim rw As Range
    For Each rw In mergedRange.Rows
        j = j + rw.RowHeight
    Next rw

    ' �վ��e����d�򪺤ؤo
    With Selection.ShapeRange
        .LockAspectRatio = False  ' ���\�e����ҧ���
        .Height = i * 0.99    '����
        .Width = j * 0.96 ' �e��
    End With
End Sub

Sub �վ�Ӥ����e_��()
Attribute �վ�Ӥ����e_��.VB_ProcData.VB_Invoke_Func = "e\n14"

' �ֳt��: Ctrl+e
    Dim ws As Worksheet
    Set ws = ActiveSheet ' �]�w����e�u�@��
    
    Selection.Placement = xlMoveAndSize '�Ӥ���m�j�p�H�x�s�����

    ' ����X���x�s�檺�d��
    Dim mergedRange As Range
    Set mergedRange = ws.Range("B5:I20")

    ' �p��X���x�s�檺��e�`�M
    Dim i As Double
    Dim col As Range
    For Each col In mergedRange.Columns
        i = i + col.Width
    Next col

    ' �p��X���x�s�檺�C���`�M
    Dim j As Double
    Dim rw As Range
    For Each rw In mergedRange.Rows
        j = j + rw.RowHeight
    Next rw

    ' �վ��e����d�򪺤ؤo
    With Selection.ShapeRange
        .LockAspectRatio = False  ' ���\�e����ҧ���
        .Height = j * 0.96    '����
        .Width = i * 0.99 ' �e��
    End With
End Sub


Sub �R������ï�Ҧ��Ӥ�()
'
' �R333 ����
'

'
    ActiveSheet.DrawingObjects.Select
    Selection.Delete
End Sub

Sub �ƻs�W�@����U�@��()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' �]�w����e�u�@��

    ' ��� B ��̫ᦳ�ƭȪ��x�s��
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, "B").End(xlUp)

    ' �T�{�̫�@�Ӧ��ƭȪ��x�s��b B ��
    If Not lastCell Is Nothing Then
        ' �p��ƻs�d�򪺰_�l��
        Dim startRow As Long
        startRow = Int((lastCell.Row - 1) / 40) * 40 + 1
        
        ' �p��ؼнd�򪺰_�l��
        Dim destinationRow As Long
        destinationRow = startRow + 40

        ' �w�q�ӷ��d��M�ؼнd��
        Dim sourceRange As Range
        Set sourceRange = ws.Range("A" & startRow & ":I" & (startRow + 39))

        Dim destinationRange As Range
        Set destinationRange = ws.Range("A" & destinationRow & ":I" & (destinationRow + 39))

        ' �ƻs���e�M�榡
        sourceRange.Copy Destination:=destinationRange
        Application.CutCopyMode = False
    End If
End Sub



