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
    
    ' 選擇資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 檔案的資料夾"
        .Show
        
        If .SelectedItems.Count = 0 Then
            MsgBox "未選擇資料夾！", vbExclamation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
    End With
    
    ' 設置錯誤處理
    On Error Resume Next
    
    ' 獲取資料夾中的所有 Excel 檔案
    fileName = Dir(folderPath & "\*.xls*")
    
    ' 處理每個 Excel 檔案
    Do While fileName <> ""
        ' 開啟工作簿
        
        ' 一還沒壓縮過
        hasCompress = False
        
        
        Set wb = Workbooks.Open(folderPath & "\" & fileName)
        Call SelectAllAndCancelGroup
        
        ' 處理每個工作表
        For Each ws In wb.Worksheets
            ' 檢查工作表是否有圖片
            hasImages = False
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then
                    hasImages = True
                    Exit For
                End If
            Next shp
            
            ' 如果有圖片，選擇第一張圖片並打開壓縮選項對話框
            If hasImages Then
                
                For Each shp In ws.Shapes
                    If shp.Type = msoPicture Then
                        shp.Select
                        ' 顯示提示訊息
                        MsgBox "在工作表 """ & ws.Name & """ 中發現圖片。" & vbNewLine & _
                               "請在接下來的對話框中選擇您想要的壓縮選項。", vbInformation
                        ' 執行壓縮圖片命令
                        Application.CommandBars.ExecuteMso "PicturesCompress"
                        hasCompress = True
                  
                        Exit For
                    End If
                Next shp
            End If
        Next ws
        
        ' 儲存並關閉工作簿
        wb.Save
        wb.Close
        
        ' 獲取下一個檔案名
        fileName = Dir()
    Loop
    
    ' 關閉錯誤處理
    On Error GoTo 0
    
    ' 完成提示
    MsgBox "所有檔案已處理完成！", vbInformation
End Sub

Sub CompressImagesInCurrentWorkbook()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim shp As Shape
    Dim hasImages As Boolean
    
    ' 處理當前工作簿中的每個工作表
    For Each ws In wb.Worksheets
        ' 檢查工作表是否有圖片
        hasImages = False
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Then
                hasImages = True
                Exit For
            End If
        Next shp
        
        ' 如果有圖片，選擇第一張圖片並打開壓縮選項對話框
        If hasImages Then
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then
                    shp.Select
                    ' 顯示提示訊息
                    MsgBox "在工作表 """ & ws.Name & """ 中發現圖片。" & vbNewLine & _
                           "請在接下來的對話框中選擇您想要的壓縮選項。", vbInformation
                    ' 執行壓縮圖片命令
                    Application.CommandBars.ExecuteMso "PicturesCompress"
                    Exit For
                End If
            Next shp
        End If
    Next ws
    
    ' 完成提示
    MsgBox "當前工作簿已處理完成！", vbInformation
End Sub



Sub SelectAllAndCancelGroup()
    ' 先選取所有工作表
    Worksheets.Select
    ' 再取消群組選取，僅保留第一個工作表被選取
    Worksheets(1).Select
End Sub
