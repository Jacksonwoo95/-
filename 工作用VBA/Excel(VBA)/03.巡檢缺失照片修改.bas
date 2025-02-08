Attribute VB_Name = "巡檢缺失"
Sub 主程序()
Attribute 主程序.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' 執行照片調整
    Call 照片調整
    
    ' 執行長邊方向判斷和選擇
    Call 判斷多個圖形長邊方向
    
    ' 執行第二次照片調整
    ' Call 照片調整_2
End Sub

Sub 照片調整()
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

Sub 判斷多個圖形長邊方向()
    Dim shp As Shape
    Dim 長邊角度 As Double
    Dim 旋轉角度 As Double
    Dim 結果 As String
    Dim 保留選擇的物件 As New Collection
    
    結果 = ""
    
    If ActiveSheet.Shapes.Count > 0 Then
        ' 首先選擇所有物件
        ActiveSheet.Shapes.SelectAll
        
        For Each shp In Selection.ShapeRange
            旋轉角度 = (shp.Rotation Mod 360 + 360) Mod 360
            
            If shp.Width >= shp.Height Then
                長邊角度 = 旋轉角度
            Else
                長邊角度 = (旋轉角度 + 90) Mod 360
            End If
            
            結果 = 結果 & "圖形 " & shp.name & " 的長邊面向"
            
            If (長邊角度 >= 315 Or 長邊角度 <= 45) Then
                結果 = 結果 & "右。" & vbNewLine
                保留選擇的物件.Add shp
            ElseIf (長邊角度 > 45 And 長邊角度 <= 135) Then
                結果 = 結果 & "下。" & vbNewLine
            ElseIf (長邊角度 > 135 And 長邊角度 <= 225) Then
                結果 = 結果 & "左。" & vbNewLine
                保留選擇的物件.Add shp
            ElseIf (長邊角度 > 225 And 長邊角度 < 315) Then
                結果 = 結果 & "上。" & vbNewLine
            End If
        Next shp
        
        ' MsgBox 結果
        
        ' 取消所有選擇
        ActiveSheet.Shapes.SelectAll
        Selection.ShapeRange.Select False
        
        ' 重新選擇不是朝上或朝下的物件
        If 保留選擇的物件.Count > 0 Then
            保留選擇的物件(1).Select
            For i = 2 To 保留選擇的物件.Count
                保留選擇的物件(i).Select False
            Next i
            ' MsgBox "已選擇 " & 保留選擇的物件.Count & " 個不朝上或朝下的物件。"
        Else
            ' MsgBox "沒有找到不朝上或朝下的物件。"
        End If
    Else
        ' MsgBox "工作表中沒有物件。"
    End If
End Sub

Sub 照片調整_2()
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

Sub 輸出缺失文字()
    Dim result As String
    Dim lastRow As Long
    Dim i As Long
    Dim lineResult As String
    
    ' 清空 M1 儲存格
    Range("M1").Value = ""
    
    ' 找出最後一個需要檢查的行
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' 檢查並連接指定儲存格的值
    For i = 1 To lastRow Step 23  ' 每 23 行檢查一次
        lineResult = ConcatIfNotEmpty(Range("A" & i)) & _
                     ConcatIfNotEmpty(Range("F" & i)) & _
                     ConcatIfNotEmpty(Range("A" & (i + 1))) & _
                     ConcatIfNotEmpty(Range("F" & (i + 1)))
        
        ' 如果這一組有任何非空值，添加到結果中並換行
        If Len(lineResult) > 0 Then
            result = result & Left(lineResult, Len(lineResult) - 1) & vbNewLine
        End If
    Next i
    
    ' 移除結果字串末尾的換行符（如果有的話）
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If
    
    ' 將結果寫入 M1 儲存格
    Range("M1").Value = result
    
    ' 設置儲存格格式為自動換行
    Range("M1").WrapText = True
End Sub

Function ConcatIfNotEmpty(cell As Range) As String
    If Not IsEmpty(cell) Then
        ConcatIfNotEmpty = cell.Value & " "
    Else
        ConcatIfNotEmpty = ""
    End If
End Function


