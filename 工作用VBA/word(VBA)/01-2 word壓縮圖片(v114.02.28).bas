Attribute VB_Name = "word壓縮圖片"
Sub WordClaude3CompressImagesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim shp As InlineShape
    Dim hasImages As Boolean
    Dim compressionDone As Boolean
    
    ' 選擇資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Word 檔案的資料夾"
        .Show
        
        If .SelectedItems.Count = 0 Then
            MsgBox "未選擇資料夾！", vbExclamation
            Exit Sub
        End If
        
        folderPath = .SelectedItems(1)
    End With
    
    ' 設置錯誤處理
    On Error Resume Next
    
    ' 獲取資料夾中的所有 Word 檔案
    fileName = Dir(folderPath & "\*.doc*")
    
    ' 處理每個 Word 檔案
    Do While fileName <> ""
        ' 開啟文件
        Set doc = Documents.Open(folderPath & "\" & fileName)
        compressionDone = False ' 重設壓縮標記
        
        ' 檢查是否有圖片
        hasImages = False
        For Each shp In doc.InlineShapes
            If shp.Type = wdInlineShapePicture Then
                hasImages = True
                Exit For
            End If
        Next shp
        
        ' 如果有圖片且尚未執行壓縮
        If hasImages And Not compressionDone Then
            For Each shp In doc.InlineShapes
                If shp.Type = wdInlineShapePicture Then
                    ' 選擇圖片
                    shp.Select
                    ' 顯示提示訊息
                    MsgBox "在檔案 """ & fileName & """ 中發現圖片。" & vbNewLine & _
                           "請在接下來的對話框中選擇您想要的壓縮選項。" & vbNewLine & _
                           "此設定將套用於此檔案中的所有圖片。", vbInformation
                    ' 執行壓縮圖片命令
                    CommandBars.ExecuteMso "PicturesCompress"
                    compressionDone = True
                    Exit For
                End If
            Next shp
        End If
        
        ' 儲存並關閉文件
        doc.Save
        doc.Close
        
        ' 獲取下一個檔案名
        fileName = Dir()
    Loop
    
    ' 關閉錯誤處理
    On Error GoTo 0
    
    ' 完成提示
    MsgBox "所有檔案已處理完成！", vbInformation
End Sub

Sub CompressImagesInCurrentDocument()
    Dim shp As InlineShape
    Dim hasImages As Boolean
    Dim compressionDone As Boolean
    
    compressionDone = False ' 初始化壓縮標記
    
    ' 檢查是否有圖片
    hasImages = False
    For Each shp In ActiveDocument.InlineShapes
        If shp.Type = wdInlineShapePicture Then
            hasImages = True
            Exit For
        End If
    Next shp
    
    ' 如果有圖片且尚未執行壓縮
    If hasImages And Not compressionDone Then
        For Each shp In ActiveDocument.InlineShapes
            If shp.Type = wdInlineShapePicture Then
                ' 選擇圖片
                shp.Select
                ' 顯示提示訊息
                MsgBox "在當前文件中發現圖片。" & vbNewLine & _
                       "請在接下來的對話框中選擇您想要的壓縮選項。" & vbNewLine & _
                       "此設定將套用於文件中的所有圖片。", vbInformation
                ' 執行壓縮圖片命令
                CommandBars.ExecuteMso "PicturesCompress"
                compressionDone = True
                Exit For
            End If
        Next shp
    End If
    
    ' 完成提示
    MsgBox "當前文件已處理完成！", vbInformation
End Sub



