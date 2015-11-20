VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CdataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim thisSt As Worksheet
Dim headerLine As Long '表头所在行，默认下一行是数据起始行

Private Sub Class_Initialize()
    Set thisSt = ActiveSheet    '默认以当前激活的表格为关联表
    headerLine = 1          '默认第一行是表头
End Sub

Sub SetUp( _
    ByVal st_ As Worksheet, _
    Optional ByVal headerLine_ As Long = 1)
    
    Set thisSt = st_
    headerLine = headerLine_
    
End Sub

'判断另一张表的表头是否跟thisSt相同（允许theOtherSt末尾多出不同列）
Function isInTheSameHeader(ByVal theOtherSt As Worksheet) As Boolean
    Dim j As Long
    For j = 1 To thisSt.UsedRange.Columns.Count
        If thisSt.Cells(headerLine, j) <> theOtherSt.Cells(headerLine, j) Then
            isInTheSameHeader = False
            Stop
            Exit Function
        End If
    Next j
    isInTheSameHeader = True
End Function

Sub ClearData()
    Dim dataStartLine As Long: dataStartLine = headerLine + 1
    Dim dataEndLine As Long: dataEndLine = thisSt.UsedRange.Rows.Count
    If dataEndLine > dataStartLine Then thisSt.Rows(dataStartLine & ":" & dataEndLine).Clear
    Debug.Print "对 "; thisSt.name; " 删除了"; dataEndLine - headerLine; " 条记录"
End Sub

Sub AddData(ByVal fromSt As Worksheet, Optional ByVal paste As XlPasteType = xlPasteAll)
'(1)先检查表头是否一致
    If Not isInTheSameHeader(fromSt) Then
        Debug.Print "thisSt:", thisSt.name; " fromSt:", fromSt.name; " 表头有差，不进行数据读取"
    End If
'(2)计算参数值
    Dim lastLine1 As Long: lastLine1 = thisSt.UsedRange.Rows.Count
    Dim lastLine2 As Long: lastLine2 = fromSt.UsedRange.Rows.Count
    Dim n As Long: n = lastLine2 - headerLine   '要增加的数据量
'(3)拷贝
    If n <> 0 Then
        fromSt.Rows((headerLine + 1) & ":" & (headerLine + n)).Copy
        thisSt.Rows((lastLine1 + 1) & ":" & (lastLine1 + n)).PasteSpecial paste
    End If
    Debug.Print "从 "; fromSt.name; " 拷贝了"; n; "条记录到 "; thisSt.name
End Sub
