Attribute VB_Name = "Module1"
Sub �]�w�ؤo()
Attribute �]�w�ؤo.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim x As Single  ' �w�q�����ܼ� x
    Dim y As Single  ' �w�q�e���ܼ� y

    ' ���ܼƽ�ȡA�o�̥H�ܨҭȽ��
    x = 10.5  ' ���]���׬�10�]���i�H�O�����B�I���A���M��z���ݨD�^
    y = 13   ' ���]�e�׬�9

    ' ���U�Ӫ��N�X�i�H�ϥ� x �M y �ܼ�
    ' �Ҧp�A�վ�襤�ϧΪ��ؤo
    With Selection.ShapeRange
        .LockAspectRatio = msoFalse
        .Height = x * 28.35 ' ���]���O�����A�ഫ���I
        .Width = y * 28.35  ' ���]���O�����A�ഫ���I
    End With
End Sub

Sub �ϳ]�w�ؤo()
    Dim x As Single  ' �w�q�����ܼ� x
    Dim y As Single  ' �w�q�e���ܼ� y

    ' ���ܼƽ�ȡA�o�̥H�ܨҭȽ��
    y = 10.5  ' ���]���׬�10�]���i�H�O�����B�I���A���M��z���ݨD�^
    x = 13   ' ���]�e�׬�9

    ' ���U�Ӫ��N�X�i�H�ϥ� x �M y �ܼ�
    ' �Ҧp�A�վ�襤�ϧΪ��ؤo
    With Selection.ShapeRange
        .LockAspectRatio = msoFalse
        .Height = x * 28.35 ' ���]���O�����A�ഫ���I
        .Width = y * 28.35  ' ���]���O�����A�ഫ���I
    End With
End Sub


Sub �R333()
'
' �R333 ����
'

'
    ActiveSheet.DrawingObjects.Select
    Selection.Delete
End Sub
