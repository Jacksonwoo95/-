Attribute VB_Name = "Module1"
Sub 設定尺寸()
Attribute 設定尺寸.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim x As Single  ' 定義長度變數 x
    Dim y As Single  ' 定義寬度變數 y

    ' 為變數賦值，這裡以示例值賦值
    x = 10.5  ' 假設長度為10（單位可以是公分、點等，取決於您的需求）
    y = 13   ' 假設寬度為9

    ' 接下來的代碼可以使用 x 和 y 變數
    ' 例如，調整選中圖形的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = msoFalse
        .Height = x * 28.35 ' 假設單位是公分，轉換為點
        .Width = y * 28.35  ' 假設單位是公分，轉換為點
    End With
End Sub

Sub 反設定尺寸()
    Dim x As Single  ' 定義長度變數 x
    Dim y As Single  ' 定義寬度變數 y

    ' 為變數賦值，這裡以示例值賦值
    y = 10.5  ' 假設長度為10（單位可以是公分、點等，取決於您的需求）
    x = 13   ' 假設寬度為9

    ' 接下來的代碼可以使用 x 和 y 變數
    ' 例如，調整選中圖形的尺寸
    With Selection.ShapeRange
        .LockAspectRatio = msoFalse
        .Height = x * 28.35 ' 假設單位是公分，轉換為點
        .Width = y * 28.35  ' 假設單位是公分，轉換為點
    End With
End Sub


Sub 刪333()
'
' 刪333 巨集
'

'
    ActiveSheet.DrawingObjects.Select
    Selection.Delete
End Sub
