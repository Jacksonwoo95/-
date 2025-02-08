Attribute VB_Name = "word���Y�Ϥ�"
Sub WordClaude3CompressImagesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim shp As InlineShape
    Dim hasImages As Boolean
    Dim compressionDone As Boolean
    
    ' ��ܸ�Ƨ�
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "�п�ܥ]�t Word �ɮת���Ƨ�"
        .Show
        
        If .SelectedItems.Count = 0 Then
            MsgBox "����ܸ�Ƨ��I", vbExclamation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
    End With
    
    ' �]�m���~�B�z
    On Error Resume Next
    
    ' �����Ƨ������Ҧ� Word �ɮ�
    fileName = Dir(folderPath & "\*.doc*")
    
    ' �B�z�C�� Word �ɮ�
    Do While fileName <> ""
        ' �}�Ҥ��
        Set doc = Documents.Open(folderPath & "\" & fileName)
        compressionDone = False ' ���]���Y�аO
        
        ' �ˬd�O�_���Ϥ�
        hasImages = False
        For Each shp In doc.InlineShapes
            If shp.Type = wdInlineShapePicture Then
                hasImages = True
                Exit For
            End If
        Next shp
        
        ' �p�G���Ϥ��B�|���������Y
        If hasImages And Not compressionDone Then
            For Each shp In doc.InlineShapes
                If shp.Type = wdInlineShapePicture Then
                    ' ��ܹϤ�
                    shp.Select
                    ' ��ܴ��ܰT��
                    MsgBox "�b�ɮ� """ & fileName & """ ���o�{�Ϥ��C" & vbNewLine & _
                           "�Цb���U�Ӫ���ܮؤ���ܱz�Q�n�����Y�ﶵ�C" & vbNewLine & _
                           "���]�w�N�M�Ω��ɮפ����Ҧ��Ϥ��C", vbInformation
                    ' �������Y�Ϥ��R�O
                    CommandBars.ExecuteMso "PicturesCompress"
                    compressionDone = True
                    Exit For
                End If
            Next shp
        End If
        
        ' �x�s���������
        doc.Save
        doc.Close
        
        ' ����U�@���ɮצW
        fileName = Dir()
    Loop
    
    ' �������~�B�z
    On Error GoTo 0
    
    ' ��������
    MsgBox "�Ҧ��ɮפw�B�z�����I", vbInformation
End Sub

Sub CompressImagesInCurrentDocument()
    Dim shp As InlineShape
    Dim hasImages As Boolean
    Dim compressionDone As Boolean
    
    compressionDone = False ' ��l�����Y�аO
    
    ' �ˬd�O�_���Ϥ�
    hasImages = False
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Then
            hasImages = True
            Exit For
        End If
    Next shp
    
    ' �p�G���Ϥ��B�|���������Y
    If hasImages And Not compressionDone Then
        For Each shp In ActiveDocument.InlineShapes
            If shp.Type = wdInlineShapePicture Then
                ' ��ܹϤ�
                shp.Select
                ' ��ܴ��ܰT��
                MsgBox "�b��e��󤤵o�{�Ϥ��C" & vbNewLine & _
                       "�Цb���U�Ӫ���ܮؤ���ܱz�Q�n�����Y�ﶵ�C" & vbNewLine & _
                       "���]�w�N�M�Ω��󤤪��Ҧ��Ϥ��C", vbInformation
                ' �������Y�Ϥ��R�O
                CommandBars.ExecuteMso "PicturesCompress"
                compressionDone = True
                Exit For
            End If
        Next shp
    End If
    
    ' ��������
    MsgBox "��e���w�B�z�����I", vbInformation
End Sub



