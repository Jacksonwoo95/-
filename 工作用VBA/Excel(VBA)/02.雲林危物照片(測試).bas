Attribute VB_Name = "雲林危物VBA1"
Sub 複製第一頁格式連續3次()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 設定為當前工作表

    Dim startRow As Long
    startRow = 1 ' 從第1行開始

    Dim copyRange As Range
    Set copyRange = ws.Range("A1:I40")

    Dim loopCounter As Integer
    loopCounter = 0 ' 計數器初始值

    ' 繼續複製直到達到 10 次或沒有足夠的行來複製內容
    Do While loopCounter < 3 And startRow <= ws.Rows.Count - 80 ' 確保有足夠的行來複製內容
        copyRange.Copy Destination:=ws.Range("A" & startRow + 40)
        startRow = startRow + 40
        loopCounter = loopCounter + 1
    Loop
End Sub

Sub 調整照片長寬()
Attribute 調整照片長寬.VB_ProcData.VB_Invoke_Func = "q\n14"

' 快速鍵: Ctrl+q
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 設定為當前工作表
    
    
    Selection.Placement = xlMoveAndSize '照片位置大小隨儲存格而變
    
    ' 獲取合併儲存格的範圍
    Dim mergedRange As Range
    Set mergedRange = ws.Range("B5:I20")

    ' 計算合併儲存格的欄寬總和
    Dim i As Double
    Dim col As Range
    For Each col In mergedRange.Columns
        i = i + col.Width
    Next col

    ' 計算合併儲存格的列高總和
    Dim j As Double
    Dim rw As Range
    For Each rw In mergedRange.Rows
        j = j + rw.RowHeight
    Next rw

    ' 調整當前選取範圍的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = False  ' 允許寬高比例改變
        .Height = i * 0.99    '高度
        .Width = j * 0.96 ' 寬度
    End With
End Sub

Sub 調整照片長寬_反()
Attribute 調整照片長寬_反.VB_ProcData.VB_Invoke_Func = "e\n14"

' 快速鍵: Ctrl+e
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 設定為當前工作表
    
    Selection.Placement = xlMoveAndSize '照片位置大小隨儲存格而變

    ' 獲取合併儲存格的範圍
    Dim mergedRange As Range
    Set mergedRange = ws.Range("B5:I20")

    ' 計算合併儲存格的欄寬總和
    Dim i As Double
    Dim col As Range
    For Each col In mergedRange.Columns
        i = i + col.Width
    Next col

    ' 計算合併儲存格的列高總和
    Dim j As Double
    Dim rw As Range
    For Each rw In mergedRange.Rows
        j = j + rw.RowHeight
    Next rw

    ' 調整當前選取範圍的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = False  ' 允許寬高比例改變
        .Height = j * 0.96    '高度
        .Width = i * 0.99 ' 寬度
    End With
End Sub


Sub 刪除活頁簿所有照片()
'
' 刪333 巨集
'

'
    ActiveSheet.DrawingObjects.Select
    Selection.Delete
End Sub

Sub 複製上一頁到下一頁()
    Dim ws As Worksheet
    Set ws = ActiveSheet ' 設定為當前工作表

    ' 找到 B 欄最後有數值的儲存格
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, "B").End(xlUp)

    ' 確認最後一個有數值的儲存格在 B 欄
    If Not lastCell Is Nothing Then
        ' 計算複製範圍的起始行
        Dim startRow As Long
        startRow = Int((lastCell.Row - 1) / 40) * 40 + 1
        
        ' 計算目標範圍的起始行
        Dim destinationRow As Long
        destinationRow = startRow + 40

        ' 定義來源範圍和目標範圍
        Dim sourceRange As Range
        Set sourceRange = ws.Range("A" & startRow & ":I" & (startRow + 39))

        Dim destinationRange As Range
        Set destinationRange = ws.Range("A" & destinationRow & ":I" & (destinationRow + 39))

        ' 複製內容和格式
        sourceRange.Copy Destination:=destinationRange
        Application.CutCopyMode = False
    End If
End Sub



