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
    
    ' 設定要處理的資料夾路徑
    ' 請修改為你的資料夾路徑
    folderPath = "C:\Users\gja552\Desktop\excel壓縮測試\消防"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' 優化執行效率
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error Resume Next ' 忽略無法處理的文件
    
    For Each file In folder.Files
        ' 檢查文件擴展名
        ext = fso.GetExtensionName(file.Name)
        If LCase(ext) = "xlsx" Or LCase(ext) = "xlsm" Or LCase(ext) = "xls" Then
            ' 開啟文件
            Set wb = Workbooks.Open(file.Path)
            
            ' 遍歷所有工作表
            For Each ws In wb.Worksheets
                ' 遍歷所有圖形物件
                For Each shp In ws.Shapes
                    ' 檢查是否為圖片類型
                    If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                        ' 執行圖片壓縮(96ppi對應msoTargetScreen)
                        shp.PictureFormat.Compress msoTargetScreen, msoPictureColorModeAutomatic
                    End If
                Next shp
            Next ws
            
            ' 保存並關閉文件
            wb.Close SaveChanges:=True
        End If
    Next file
    
    ' 恢復原始設定
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    MsgBox "所有圖片已壓縮完成!", vbInformation
End Sub
