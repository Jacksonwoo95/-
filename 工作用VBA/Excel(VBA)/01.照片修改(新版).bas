Attribute VB_Name = "測試照片"
Sub 照片調整() '照片依當前欄寬列高調整
Attribute 照片調整.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' 照片調整 巨集
'
' 快速鍵: Ctrl+q
'
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range

    Set ws = ActiveSheet  ' 使用當前活動的工作表
    Set mergedRange = ws.Range("A1:E24")  ' 設定為合併的範圍

    ' 獲取合併儲存格的總長度（高度）和總寬度
    i = mergedRange.Height  ' 儲存格的總長度（即高度）
    j = mergedRange.Width  ' 儲存格的總寬度

    ' 將值輸入 L1 和 M1 儲存格
    'ws.Range("L1").Value = i
    'ws.Range("M1").Value = j

    ' 調整當前選取範圍的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = False  ' 允許寬高比例改變
        .Height = j '寬度
        .Width = i * 0.95   ' 0.95倍高度
    End With
End Sub
Sub 照片調整_反() '照片依當前欄寬列高調整(長寬對調)
Attribute 照片調整_反.VB_ProcData.VB_Invoke_Func = "e\n14"

'
' 照片反向 巨集
'
' 快速鍵: Ctrl+e
'
    Dim i As Double, j As Double
    Dim ws As Worksheet
    Dim mergedRange As Range

    Set ws = ActiveSheet  ' 使用當前活動的工作表
    Set mergedRange = ws.Range("A1:E24")  ' 設定為合併的範圍

    ' 獲取合併儲存格的總長度（高度）和總寬度
    i = mergedRange.Height  ' 儲存格的總長度（即高度）
    j = mergedRange.Width  ' 儲存格的總寬度

    ' 將值輸入 L1 和 M1 儲存格
    'ws.Range("L1").Value = i
    'ws.Range("M1").Value = j

    ' 調整當前選取範圍的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = False  ' 允許寬高比例改變
        .Height = i * 0.95    ' 0.95倍高度
        .Width = j ' 寬度
    End With
    
End Sub


Sub 旋轉()
    Dim i As Integer
    For i = 1 To 3
        Selection.ShapeRange.IncrementRotation 90
    Next i
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

Sub SetAlternatingBorderColors()
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


