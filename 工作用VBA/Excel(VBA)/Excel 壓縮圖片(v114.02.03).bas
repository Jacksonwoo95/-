Attribute VB_Name = "���Y�Ϥ�"
Sub ���Y�Ϥ�����()
Attribute ���Y�Ϥ�����.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���Y�Ϥ����� ����
'

        Application.CommandBars.ExecuteMso "PicturesCompress"

End Sub

Sub BatchCompressPictures()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim folderPath As String

    ' �]�w�s�� Excel �ɮת���Ƨ����|�A�Цۦ�קאּ��ڸ��|
    folderPath = "C:\Users\gja552\Desktop\excel���Y����\����"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' �v�@�B�z��Ƨ������C�� Excel �ɮ�
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
            Set wb = Workbooks.Open(file.Path)
            
            ' ���ɮפ����C�Ӥu�@��i��B�z
            For Each ws In wb.Worksheets
                On Error Resume Next
                ' ���տ���Ҧ��Ϥ��]Shapes�^
                ws.Shapes.SelectAll
                ' �I�s���Y�Ϥ������ث��O�]�`�N�G�o�өR�O�i��]�����Ӳ��A�Y���ѥi�H��Ψ�L��k�ζi����~�B�z�^
                Application.CommandBars.ExecuteMso "PicturesCompress"
                On Error GoTo 0
            Next ws
            
            wb.Close SaveChanges:=True
        End If
    Next file
    
    MsgBox "�妸�B�z�����I"
End Sub

