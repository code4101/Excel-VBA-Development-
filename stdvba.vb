'code4101 最新更新于:2015年5月20日 6:40
' 10:31 修正“is该表存在”函数的一个低级bug
' 11:50 增加颜色枚举常量表

' 5月20日
' 3:36: 加入工具箱里的CZ函数（查找）
' 5:05: 增加 保留区域 过程
' 5:34: 加入 自动填充 过程
' 6:40: 添加 转数值 函数

' 5月26日
' 10:02: 自动填充过程的参数类型改成byval

' 5月27日
' 16:21:增加 函数：该列使用的最后一个单元格，过程：单列智能拷贝

'8月2日
'加入表头查找系列工具
'删除“区域的交”函数，因为可以用VBA自带的Intersect代替
'加入addr、openFile函数
'加入hasStn：判断工作薄是否存在名称为stn的表格

'8月3日
'openFile删除（因为用Workbooks.open方法已经很方便），改为"优化路径"的getfn


Enum 颜色表
    标准字段颜色 = 15773696 'RGB(0, 176, 240)   蓝色
    额外字段颜色 = 5296274  'RGB(146, 208, 80)  绿色
    字段分类颜色 = 65535    'RGB(255, 255, 0)   黄色
End Enum


Function 该列使用的最后一个单元格(ByVal x As Range) As Range
''算法思路: 使用三个指针，y指向x的下一个跳跃点，z指向y的下一个跳跃点
'' 当y与z所指相同时，此时x即为内容结尾
'    Dim y As Object, z As Object
'    Set y = x.End(xlDown)
'    Set z = y.End(xlDown)
'    Do While y.Address <> z.Address
'        Set x = y
'        Set y = z
'        Set z = z.End(xlDown)
'    Loop
'    Set 该列使用的最后一个单元格 = x
'2015/7/29日更新
    '首先更改x的参数类型object为range
    '然后改变算法
    Dim st As Worksheet: Set st = x.Parent
    With st
        Set 该列使用的最后一个单元格 = .Cells(.Cells(.Rows.Count, x.Column).End(xlUp).Row, x.Column)
    End With
End Function

'将x至x末尾所在列的内容拷贝到y单元格及后面
Sub 单列智能拷贝(x As Range, y As Range)
    Range(x, 该列使用的最后一个单元格(x)).Copy y
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''sheet等相关信息提取''''''''''''''''''''''''''''''''''''''''''''
' 从(i,j)开始，往下遍历，遇到空代表最后一行
Function 最后一行(一张表, Optional ByVal i As Integer = 1, Optional ByVal j As Integer = 1)
    Do While 一张表.Cells(i, j) <> ""
        i = i + 1
    Loop
    最后一行 = i - 1
End Function


' 从(i,j)开始，往右遍历，遇到空代表最后一列
Public Function 最后一列(一张表, Optional ByVal i As Integer = 1, Optional ByVal j As Integer = 1)
    Do While 一张表.Cells(i, j) <> ""
        j = j + 1
    Loop
    最后一列 = j - 1
End Function

' 在(x1,y1)至(x2,y2)内找指定内容
' 若找到，则将结果存储在x1,y1
' 返回布尔值:true代表成功,false代表失败
Function 在一定范围内查找指定文本所在位置(一张表, 查找值, ByRef x1, ByRef y1, x2 As Integer, y2 As Integer) As Boolean
    For i = x1 To x2
        For j = y1 To y2
            If 一张表.Cells(i, j) = 查找值 Then
                x1 = i
                y1 = j
                在一定范围内查找指定文本所在位置 = True
                Exit Function
            End If
        Next j
    Next i
    在一定范围内查找指定文本所在位置 = False
End Function

'使用举例:  Debug.Print is该表存在(Workbooks("电信提取表.xlsb"), "Sheet1")
Function is该表存在(工作薄, 表名 As String) As Boolean
    is该表存在 = False
    For i = 1 To 工作薄.Sheets.Count
        If 表名 = 工作薄.Sheets(i).name Then
            is该表存在 = True
            Exit Function
        End If
    Next i
End Function



' 删除第i行，第j列外的单元格
Sub 保留区域(i As Integer, j As String)
    Rows(i & ":1048576").Delete Shift:=xlUp
    Columns(j & ":XFD").Delete Shift:=xlToLeft
End Sub

' 按照第一行的情况，自动填充至(x2,y2)
Sub 自动填充(ByVal x1 As Integer, ByVal y1, ByVal x2 As Integer, ByVal y2)
    Range(Cells(x1, y1), Cells(x1, y2)).Select
    Selection.AutoFill Destination:=Range(Cells(x1, y1), Cells(x2, y2)), Type:=xlFillDefault
End Sub

Function 转数值(a) As Double
    If a = "" Then
        转数值 = 0#
    Else
        转数值 = a
    End If
End Function



'''''''''''''''''''''''''''''''''''''''''''''''数学库(math.h)'''''''''''''''''''''''''''''''''''''''''''''
Function pow(a, m)
' a是方阵,m是不小于1的整数次幂
    pow = a
    m = m - 1
    While m > 0
        If m Mod 2 Then
            pow = Application.WorksheetFunction.MMult(pow, a)
        End If
        a = Application.WorksheetFunction.MMult(a, a)
        m = Int(m / 2)
    Wend
End Function

' 将列的数字编号转换为字母编号
Function 列名(列号 As Integer) As String
    Do While 列号 > 25
        列号 = 列号 - 26
        列名 = 列名 + "Z"
    Loop
    If 列号 > 0 Then 列名 = 列名 + Chr(64 + 列号)
End Function

'''''''''''''''''''''''''''''''''''''''''''''算法库(algorithm)'''''''''''''''''''''''''''''''''''''''''''''
' 来源:http://stackoverflow.com/questions/152319/vba-array-sort-function
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)

  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)

     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If

  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi

End Sub

Function CZ(查找值 As String, 查找值所在区域 As Range, Optional 目标值所在列 As Variant, Optional 确认返回第几个目标值 As Integer = 1, Optional 模糊查找 As Integer = 1) As String
    Application.Volatile
    Dim i As Long, R As Range, R1 As Range, Str As String, L As Long
    Dim CZFS As Long
    Dim st As String, p As Long
    
    If 模糊查找 = 2 Then   '1：常规模糊查找，2：超级模糊查找
       st = ""
       For p = 1 To Len(查找值)
           st = st & Mid(查找值, p, 1) & "*"
       Next p
       查找值 = Left(st, Len(st) - 1)
    End If
    
    If 模糊查找 > 0 Then CZFS = xlPart Else CZFS = xlWhole
    
    Dim sh As Worksheet, SH1 As Worksheet
    
      
    With 查找值所在区域(1).Resize(查找值所在区域.Rows.Count, 1)
    If .Cells(1) = 查找值 Then Set R = .Cells(1) Else Set R = .Find(查找值, LookIn:=xlValues, lookat:=CZFS)
     If Not R Is Nothing Then
        Set sh = R.Parent
     
     
        If TypeName(目标值所在列) = "Range" Then
           Set R1 = 目标值所在列
           Set SH1 = R1.Parent
           L = 目标值所在列.Column
        Else
           L = 目标值所在列
           If L = 0 Then L = R.Column
        End If
     
        Str = R.Address
        Do
            i = i + 1
            If i = 确认返回第几个目标值 Then
              If Not SH1 Is Nothing Then CZ = SH1.Cells(R.Row, L) Else CZ = Cells(R.Row, L)
              Exit Function
            End If
            Set R = 查找值所在区域.Find(查找值, R, lookat:=CZFS)
        Loop While Not R Is Nothing And R.Address <> Str
    End If
End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''字符串库(algorithm)'''''''''''''''''''''''''''''''''''''''''''''
Function 字符串相似度(全名 As String, 简称 As String)
    长度1 = Len(全名)
    长度2 = Len(简称)
    
    字符串相似度 = 0
    k = 1
    For i = 1 To 长度2
            
        ' 在 k~长度1 找 mid(简称,i,1), 记位置为j
        pos = -1
        For j = k To 长度1
            If Mid(简称, i, 1) = Mid(全名, j, 1) Then
                pos = j
                Exit For
            End If
        Next j
        
        If pos <> -1 Then
            字符串相似度 = 字符串相似度 + 1
            k = pos
        End If
        
    Next i
    
End Function


Function 字符连接(rng As Range, Optional 行分隔符 As String = ",", Optional 列分隔符 As String = ";") As String

    For i = 1 To rng.Rows.Count
        
        If i <> 1 Then 字符连接 = 字符连接 & 列分隔符
        字符连接 = 字符连接 & rng.Cells(i, 1)
        
        For j = 2 To rng.Columns.Count
            字符连接 = 字符连接 & 行分隔符 & rng.Cells(i, j)
        Next j
        
    Next i
    
End Function


' 输入的x和y都是一个单元格
' maxn可选参数表示最多映射的键值数
Function 一对多的值汇总(x, y, Optional maxn As Integer = 20, Optional 分隔符 As String = ",")
    '(0)如果键值为空,则返回空
    If x = "" Then
        一对多的值汇总 = ""
        Exit Function
    End If
    '(1) 计算总行数
    Dim i
    For i = 1 To maxn
        If (x.Offset(i, 0) <> "" And x.Offset(i, 0) <> x) Or y.Offset(i, 0) = "" Then
            Exit For
        End If
    Next i
    '（2）调用子函数对范围内的值进行拼接
    一对多的值汇总 = 字符连接(Range(y, y.Offset(i - 1, 0)), "", 分隔符)
End Function

Function onlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function

Function CleanString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .Pattern = "[^\d]+"
    CleanString = .Replace(strIn, vbNullString)
    End With
End Function




'{表头查找系列工具
'''以下为7月30日
Function findcol(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Long
    Dim t As Range
    Set t = findcel(st, name, partName)
    If t Is Nothing Then
        findcol = 0
    Else
        findcol = t.Column
    End If
End Function

Function findrow(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Long
    Dim t As Range
    Set t = findcel(st, name, partName)
    If t Is Nothing Then
        findrow = 0
    Else
        findrow = t.Row
    End If
End Function


'该函数支持name、partName用分号隔开，允许按优先级进行字段名搜索的多字段查询
Function findcel(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Range
'(1)首先name绝对不能为空
    If name = "" Then Exit Function

    Dim arr1, arr2
'(2)partName可以为空，但为了后续遍历统一处理，需要先预分析下
    arr1 = Split(partName, ";")
    If isEmptyArr(arr1) Then
        ReDim arr1(1 To 1)
        arr1(1) = ""
    End If
    
'(3)开始循环遍历,只要找到第一组满足解即可
    arr2 = Split(name, ";")
    For Each a1 In arr1
        For Each a2 In arr2
            Set findcel = findcel_base(st, a2, a1)
            If Not (findcel Is Nothing) Then Exit Function
        Next a2
    Next a1
End Function

Function findcel_base(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Range
    Dim rng As Range '查找的范围
    Set rng = st.UsedRange
    
    'Debug.Print "findcel_base查找内容所在工作薄", st.Parent.name
    Dim rng2 As Range, t As Range
'(1)先定位高级表头的列范围
    If partName <> "" Then
        Set t = rng.Find(partName, lookat:=xlPart)
        '如果第一个是合并单元格，有时候会有找不到的bug
        If rng.Cells(1, 1) = partName Then Set t = rng.Cells(1, 1)
        '如果确实找不到，退出函数
        If t Is Nothing Then Exit Function
        
        '否则就是找到了，计算出找到的(合并)单元格所在列
        Set rng2 = st.Range(rng.Cells(1, t.Column), rng.Cells(st.Rows.Count, t.Offset(0, 1).Column - 1))
        Set rng = Intersect(rng, rng2)  'Range的交
    End If

'(2)然后就可以直接在rng搜索表头名了
    Set t = rng.Find(name, lookat:=xlWhole)                        '能单元格匹配找到，则按照单元格结果
    If t Is Nothing Then Set t = rng.Find(name, lookat:=xlPart)    '否则进行部分查找
    If name = rng.Cells(1, 1) Then Set t = rng.Cells(1, 1)
    
    'If Not (t Is Nothing) Then Debug.Print name & "在" & t.Address
    Set findcel_base = t
End Function

Private Function isEmptyArr(arr) As Boolean  '
    isEmptyArr = True
    For Each a In arr
        isEmptyArr = False
        Exit For
    Next a
End Function
'表头查找系列工具}


'获得单元格的位置(去掉绝对引用符)
Function addr(cell) As String
    addr = Replace(cell.Address, "$", vbNullString)
End Function


'oriStr是原始文件名或路径，optPath是参考路径
'如果oriStr已经是有效的文件名，则返回原值
'否则，将oriStr的所在目录设置到optPath
'如果是无效路径，会返回空字符串
Function getfn(oriStr As String, Optional optPath As String) As String
'参考资料:http://zhidao.baidu.com/link?url=9qQA8dJddTAGsmuPyrKpl6IQbBnxI7PNY9-os-WZhjsj2k5V4-d95nfR6GFlr8hL3uW-RCrL_St1EouTmJiX7bU5m6KQZDBQU0_VGY_31EW
    Dim fso As Object
    Dim res As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(oriStr) Then
        res = oriStr
    ElseIf fso.FileExists(optPath & "\" & oriStr) Then
        res = optPath & "\" & oriStr
    Else
        res = ""
    End If
    getfn = res
End Function


Function hasStn(ByVal wb As Workbook, ByVal stn As String) As Boolean
    hasStn = False
    For Each st In wb.Sheets
        If st.name = stn Then hasStn = True
        Exit Function
    Next st
End Function
Private Sub hasStn测试()
    'Debug.Print hasStn(ActiveWorkbook, "2-2固定资产－铁塔")
    Debug.Print is该表存在(ActiveWorkbook, "2-2固定资产－铁塔")
End Sub

'自动调整表格宽、高
Sub autoFit(st As Worksheet)
    With st.Cells
        .EntireColumn.autoFit
        .EntireRow.autoFit
    End With
End Sub


