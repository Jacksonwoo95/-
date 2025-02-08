Attribute VB_Name = "Claude3"
Sub CompressImagesInF()
older()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim hasImages As Boolean
    Dim hasCompressI As Boolean
    
    ' ��ܸ�Ƨ�
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�п�ܥ]�t Excel �ɮת���Ƨ�"
        .Show
        
        If .SelectedItems.Count = 0 Then
            MsgBox "����ܸ�Ƨ��I", vbExclamation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
    End With
    
    ' �]�m���~�B�z
    On Error Resume Next
    
    ' �����Ƨ������Ҧ� Excel �ɮ�
    fileName = Dir(folderPath & "\*.xls*")
    
    ' �B�z�C�� Excel �ɮ�
    Do While fileName <> ""
        ' �}�Ҥu�@ï
        
        ' �@�٨S���Y�L
        hasCompress = False
        
        
        Set wb = Workbooks.Open(folderPath & "\" & fileName)
        Call SelectAllAndCancelGroup
        
        ' �B�z�C�Ӥu�@��
        For Each ws In wb.Worksheets
            ' �ˬd�u�@��O�_���Ϥ�
            hasImages = False
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then
                    hasImages = True
                    Exit For
                End If
            Next shp
            
            ' �p�G���Ϥ��A��ܲĤ@�i�Ϥ��å��}���Y�ﶵ��ܮ�
            If hasImages Then
                
                For Each shp In ws.Shapes
                    If shp.Type = msoPicture Then
                        shp.Select
                        ' ��ܴ��ܰT��
                        MsgBox "�b�u�@�� """ & ws.Name & """ ���o�{�Ϥ��C" & vbNewLine & _
                               "�Цb���U�Ӫ���ܮؤ���ܱz�Q�n�����Y�ﶵ�C", vbInformation
                        ' �������Y�Ϥ��R�O
                        Application.CommandBars.ExecuteMso "PicturesCompress"
                        hasCompress = True
                  
                        Exit For
                    End If
                Next shp
            End If
        Next ws
        
        ' �x�s�������u�@ï
        wb.Save
        wb.Close
        
        ' ����U�@���ɮצW
        fileName = Dir()
    Loop
    
    ' �������~�B�z
    On Error GoTo 0
    
    ' ��������
    MsgBox "�Ҧ��ɮפw�B�z�����I", vbInformation
End Sub

Sub CompressImagesInCurrentWorkbook()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim shp As Shape
    Dim hasImages As Boolean
    
    ' �B�z��e�u�@ï�����C�Ӥu�@��
    For Each ws In wb.Worksheets
        ' �ˬd�u�@��O�_���Ϥ�
        hasImages = False
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Then
                hasImages = True
                Exit For
            End If
        Next shp
        
        ' �p�G���Ϥ��A��ܲĤ@�i�Ϥ��å��}���Y�ﶵ��ܮ�
        If hasImages Then
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then
                    shp.Select
                    ' ��ܴ��ܰT��
                    MsgBox "�b�u�@�� """ & ws.Name & """ ���o�{�Ϥ��C" & vbNewLine & _
                           "�Цb���U�Ӫ���ܮؤ���ܱz�Q�n�����Y�ﶵ�C", vbInformation
                    ' �������Y�Ϥ��R�O
                    Application.CommandBars.ExecuteMso "PicturesCompress"
                    Exit For
                End If
            Next shp
        End If
    Next ws
    
    ' ��������
    MsgBox "��e�u�@ï�w�B�z�����I", vbInformation
End Sub



Sub SelectAllAndCancelGroup()
    ' ������Ҧ��u�@��
    Worksheets.Select
    ' �A�����s�տ���A�ȫO�d�Ĥ@�Ӥu�@��Q���
    Worksheets(1).Select
End Sub
