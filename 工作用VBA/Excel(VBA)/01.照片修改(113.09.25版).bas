Attribute VB_Name = "照片調整v240925"
Sub 主程序()
Attribute 主程序.VB_ProcData.VB_Invoke_Func = "q\n14"
    ' 執行照片調整
    Call 照片調整
    
    ' 執行長邊方向判斷和選擇
    Call 判斷多個圖形長邊方向
    
End Sub

Sub 照片調整()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A1:E24")
    
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
        ' ActiveSheet.Shapes.SelectAll
        Selection.ShapeRange.Select False
        
        ' 重新選擇不是朝上或朝下的物件
        If 保留選擇的物件.Count > 0 Then
            保留選擇的物件(1).Select
            For i = 2 To 保留選擇的物件.Count
                保留選擇的物件(i).Select False
            Next i
            Call 照片調整_2
            ' MsgBox "已選擇 " & 保留選擇的物件.Count & " 個不朝上或朝下的物件。"
        Else
            ' MsgBox "沒有找到不朝上或朝下的物件。"
        End If
    Else
        MsgBox "工作表中沒有物件。"
    End If
End Sub

Sub 照片調整_2()
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range
    Set ws = ActiveSheet
    Set mergedRange = ws.Range("A1:E24")
    
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

Sub SetAlternatingBorderColors()
Attribute SetAlternatingBorderColors.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim finalRow As Long
    Dim i As Integer

    ' 設置當前激活的工作表為目標工作表
    Set ws = ActiveSheet

    ' 獲取目標工作表中的最後一行數
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    ' 計算最終行數（在最後一行的基礎上再加上52行）
    finalRow = lastRow + 208

    ' 以26行的間隔處理所有相關的行，直到最終行
    For i = 25 To finalRow Step 26
        ApplyAlternatingBorders ws.Range("L" & i & ":O" & i)
        ApplyAlternatingBorders ws.Range("Q" & i & ":T" & i)
    Next i
End Sub

Sub ApplyAlternatingBorders(rng As Range)
    Dim cell As Range
    Dim borderColor As Long
    
    For Each cell In rng
        If (cell.Column - rng.Cells(1, 1).Column) Mod 2 = 0 Then
            borderColor = vbRed ' 偶數列設為紅色
        Else
            borderColor = vbBlue ' 奇數列設為藍色
        End If
        
        ' 設置上、下、左、右的框線
        With cell.Borders
            .Item(xlEdgeTop).LineStyle = xlContinuous
            .Item(xlEdgeTop).Color = borderColor
            .Item(xlEdgeBottom).LineStyle = xlContinuous
            .Item(xlEdgeBottom).Color = borderColor
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeLeft).Color = borderColor
            .Item(xlEdgeRight).LineStyle = xlContinuous
            .Item(xlEdgeRight).Color = borderColor
        End With
    Next cell
End Sub


Sub 刪除活頁簿所有照片()
    Dim response As VbMsgBoxResult
    
    response = MsgBox("您確定要刪除所有照片嗎?", vbYesNo + vbQuestion, "確認刪除")
    
    If response = vbYes Then
        ActiveSheet.DrawingObjects.Select
        Selection.Delete
        
    End If
    ' 如果選擇 "否"，程式將直接結束，不顯示任何訊息
    
End Sub
