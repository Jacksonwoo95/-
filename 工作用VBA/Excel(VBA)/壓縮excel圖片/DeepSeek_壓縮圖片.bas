Attribute VB_Name = "DeepSeek"
Sub DeepSeekCompressAllPictures()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim folderPath As String
    Dim ext As String
    
    ' �]�w�n�B�z����Ƨ����|
    ' �Эקאּ�A����Ƨ����|
    folderPath = "C:\Users\gja552\Desktop\excel���Y����\����"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' �u�ư���Ĳv
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error Resume Next ' �����L�k�B�z�����
    
    For Each file In folder.Files
        ' �ˬd����X�i�W
        ext = fso.GetExtensionName(file.Name)
        If LCase(ext) = "xlsx" Or LCase(ext) = "xlsm" Or LCase(ext) = "xls" Then
            ' �}�Ҥ��
            Set wb = Workbooks.Open(file.Path)
            
            ' �M���Ҧ��u�@��
            For Each ws In wb.Worksheets
                ' �M���Ҧ��ϧΪ���
                For Each shp In ws.Shapes
                    ' �ˬd�O�_���Ϥ�����
                    If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                        ' ����Ϥ����Y(96ppi����msoTargetScreen)
                        shp.PictureFormat.Compress msoTargetScreen, msoPictureColorModeAutomatic
                    End If
                Next shp
            Next ws
            
            ' �O�s���������
            wb.Close SaveChanges:=True
        End If
    Next file
    
    ' ��_��l�]�w
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    MsgBox "�Ҧ��Ϥ��w���Y����!", vbInformation
End Sub
