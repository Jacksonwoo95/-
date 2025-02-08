Attribute VB_Name = "壓縮圖片"
Sub 壓縮圖片測試()
Attribute 壓縮圖片測試.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 壓縮圖片測試 巨集
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

    ' 設定存放 Excel 檔案的資料夾路徑，請自行修改為實際路徑
    folderPath = "C:\Users\gja552\Desktop\excel壓縮測試\消防"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 逐一處理資料夾中的每個 Excel 檔案
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
            Set wb = Workbooks.Open(file.Path)
            
            ' 對檔案中的每個工作表進行處理
            For Each ws In wb.Worksheets
                On Error Resume Next
                ' 嘗試選取所有圖片（Shapes）
                ws.Shapes.SelectAll
                ' 呼叫壓縮圖片的內建指令（注意：這個命令可能因版本而異，若失敗可以改用其他方法或進行錯誤處理）
                Application.CommandBars.ExecuteMso "PicturesCompress"
                On Error GoTo 0
            Next ws
            
            wb.Close SaveChanges:=True
        End If
    Next file
    
    MsgBox "批次處理完成！"
End Sub

