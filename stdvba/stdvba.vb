'code4101 最新更新于:2015年11月20日 21:00+

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
Sub 自动填充2(rng As Range)
    Dim rng2 As Range   '要填充的全部范围
    Set rng2 = rng.Parent.UsedRange '先获得所有使用范围
    
    rng.AutoFill Destination:=Range("A3:A208")
End Sub

Function 转数值(A) As Double
    If A = "" Then
        转数值 = 0#
    Else
        转数值 = A
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''数学库(math.h)'''''''''''''''''''''''''''''''''''''''''''''
Function pow(A, m)
' a是方阵,m是不小于1的整数次幂
    pow = A
    m = m - 1
    While m > 0
        If m Mod 2 Then
            pow = Application.WorksheetFunction.MMult(pow, A)
        End If
        A = Application.WorksheetFunction.MMult(A, A)
        m = Int(m / 2)
    Wend
End Function


Function 列名(ByVal 列号 As Long) As String
    If 列号 = 0 Then
        列名 = "#N/A"
        Exit Function
    End If
    Dim s As String: s = addr(Cells(1, 列号))
    列名 = Left(s, Len(s) - 1)
End Function

'' 将列的数字编号转换为字母编号
'Function 列名(列号 As Integer) As String
'    Do While 列号 > 25
'        列号 = 列号 - 26
'        列名 = 列名 + "Z"
'    Loop
'    If 列号 > 0 Then 列名 = 列名 + Chr(64 + 列号)
'End Function

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
    Dim i As Long, r As Range, R1 As Range, Str As String, L As Long
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
    If .Cells(1) = 查找值 Then Set r = .Cells(1) Else Set r = .Find(查找值, LookIn:=xlValues, lookat:=CZFS)
     If Not r Is Nothing Then
        Set sh = r.Parent
     
     
        If TypeName(目标值所在列) = "Range" Then
           Set R1 = 目标值所在列
           Set SH1 = R1.Parent
           L = 目标值所在列.Column
        Else
           L = 目标值所在列
           If L = 0 Then L = r.Column
        End If
     
        Str = r.Address
        Do
            i = i + 1
            If i = 确认返回第几个目标值 Then
              If Not SH1 Is Nothing Then CZ = SH1.Cells(r.Row, L) Else CZ = Cells(r.Row, L)
              Exit Function
            End If
            Set r = 查找值所在区域.Find(查找值, r, lookat:=CZFS)
        Loop While Not r Is Nothing And r.Address <> Str
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




'代码更新于2015年07月30日
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
    For Each A1 In arr1
        For Each A2 In arr2
            Set findcel = findcel_base(st, A2, A1)
            If Not (findcel Is Nothing) Then Exit Function
        Next A2
    Next A1
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
    If name = rng.Cells(1, 1) Then Set t = rng.Cells(1, 1)          '如果第一个单元格值满足，强制将结论修正为cells(1,1)
    '如果前面几种情况都没找到，则看cells(1,1)是否满足模糊匹配
    If t Is Nothing And rng.Cells(1, 1) Like "*" & name & "*" Then Set t = rng.Cells(1, 1)
    
    'If Not (t Is Nothing) Then Debug.Print name & "在" & t.Address
    Set findcel_base = t
End Function

Private Function isEmptyArr(arr) As Boolean  '
    isEmptyArr = True
    For Each A In arr
        isEmptyArr = False
        Exit For
    Next A
End Function


'获得单元格的位置(去掉绝对引用符)
Function addr(cell) As String
    addr = Replace(cell.Address, "$", vbNullString)
End Function


'oriStr是原始文件名或路径，optPath是参考路径
'如果oriStr已经是有效的文件名，则返回原值
'否则，将oriStr的所在目录设置到optPath
'如果是无效路径，会返回空字符串
Function getfn(ByVal oriStr As String, Optional optPath As String) As String
'参考资料:http://zhidao.baidu.com/link?url=9qQA8dJddTAGsmuPyrKpl6IQbBnxI7PNY9-os-WZhjsj2k5V4-d95nfR6GFlr8hL3uW-RCrL_St1EouTmJiX7bU5m6KQZDBQU0_VGY_31EW
    Dim fso As Object
    Dim res As String
    
    '这个fso貌似很智能，如果工作薄已打开，也能正确识别已存在
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


'is正向 为 True:程序初始化
'is正向 为 False:反初始化
Sub 程序初始化与反初始化(is正向 As Boolean)
    Static tt As Double
    If is正向 Then
        tt = Timer  '计时器
        With Application
            .ScreenUpdating = False     '关闭屏幕更新加快执行速度
            .DisplayAlerts = False
        End With
    Else
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
        End With
        MsgBox "程序运行完毕，用时 " & Format(Timer - tt, "#0.00") & "秒", Title:="计时器"
    End If
End Sub


'Function getWb(ByVal wbn) As Workbook
'    Dim st0 As Worksheet: Set st0 = ActiveSheet       '用于存储原激活工作表
'
'    Debug.Print Chr(10); "getWb函数中，尝试打开文件："; CStr(wbn)
'    wbn = getfn(CStr(wbn))
'
'    If wbn = "" Then
'        Debug.Print "但该文件路径无效"
'        Exit Function
'    Else
'        Debug.Print "该文件名有效"
'        '如果该文件已打开
'        For Each wb In Workbooks
'            If wbn Like ("*" & wb.name & "*") Then
'                Debug.Print "不过该文件已经处于打开状态"
'                Set getWb = wb
'                Exit Function
'            End If
'        Next wb
'
'        '如果未打开
'        Set getWb = Workbooks.Open(wbn, False, False)    '使用吗默认的习惯性操作打开文件
'        If getWb Is Nothing Then
'            Debug.Print "但是仍然打开失败";
'        Else
'            st0.Activate                                '激活原"激活工作表"
'            Debug.Print "打开成功:"; getWb.name
'        End If
'    End If
'
'End Function

Function getWb(ByVal wbn) As Workbook
    Dim wb As Workbook

'(1)统一为绝对路径
    wbn = CStr(wbn)
    If Mid(wbn, 2, 1) <> ":" Then wbn = ActiveWorkbook.path & "\" & wbn

'(2)与已经打开的文件对比
    Dim wb1 As Workbook
    Dim path As String, name As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.GetParentFolderName(wbn)
    name = fso.GetFileName(wbn)
    
    For Each wb1 In Workbooks
        If wb1.path = path And wb1.name = name Then
            Debug.Print "该文件已经处于打开状态:", wbn
            Set getWb = wb1
            Exit Function
        ElseIf wb1.name = name And wb1.path <> path Then
            Debug.Print "工作薄名称重复,无法打开该文件:", wbn
            Exit Function
        End If
    Next wb1

'(3)尝试打开
    On Error Resume Next
    Set wb = Workbooks.Open(wbn, False)
    
    If wb Is Nothing Then
        Debug.Print "文件名无效,该文件未打开:", wbn
    Else
        Debug.Print "成功打开文件:", wbn
        Set getWb = wb
    End If
End Function

Private Sub 测试getWb()
    Dim wb As Workbook
    Set wb = getWb("中国移动31福建02厦门明细表0803-定稿.xlsx")
    
    Debug.Print "当前激活的工作薄路径:", ActiveWorkbook.path
    If wb Is Nothing Then
        Debug.Print "综合测试结果:未打开文件"
    Else
        Debug.Print "综合测试结果:打开文件", wb.path, wb.name
    End If
End Sub

Private Sub fso特殊测试()
    s = ActiveWorkbook.path & "\aaa"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print fso.GetParentFolderName(s)
    Debug.Print fso.GetFileName(s)
End Sub

Function has汉字(s As String) As Boolean
    has汉字 = False
    Dim t As Long, i As Long
    
    For i = 1 To Len(s)
        t = Asc(Mid(s, i, 1))
        If t < 1 Or t > 127 Then
            has汉字 = True
            Exit Function
        End If
    Next i
End Function


Function get集合Text( _
    ByVal c As Collection, _
    Optional 排序 As Boolean = True, _
    Optional 重复项计数 As Boolean = True, _
    Optional 项分隔符 As String = ", ", _
    Optional 重复项计数分隔符 = "") As String
'(1)
    If 重复项计数 Then
        Set c = 对Collection重复项汇总(c, 重复项计数分隔符)
    End If
'(2)
    If 排序 Then
        Set c = 对Collection排序(c)
    End If
'(3)
    Dim ans As String
    For Each cc In c
        If ans = "" Then
            ans = cc
        Else
            ans = ans & 项分隔符 & cc
        End If
    Next cc
'(4)return
    get集合Text = ans
End Function
Function 对Collection排序(c As Collection) As Collection
'实现原理：先将集合转成数组Array，对Array使用quicksort，然后再将排好序的Array存回新的Colllection
    Set 对Collection排序 = New Collection
    Dim A() As Variant
    A = CollectionToArray(c)
    Call QuickSort(A, LBound(A), UBound(A))
    For Each aa In A
        对Collection排序.add aa
    Next aa
End Function
Function 对Collection重复项汇总(c As Collection, Optional 重复项计数分隔符 = "") As Collection
'(1)先计算出每一项的数量
    Dim cnt As Object: Set cnt = CreateObject("Scripting.Dictionary") '用于建立字典辅助
    For Each k In c
        cnt(k) = cnt(k) + 1
    Next k
'(2)算出新的集合
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set 对Collection重复项汇总 = New Collection
    For Each k In c
        If Not d.Exists(k) Then 'd存储已经visited的项
            对Collection重复项汇总.add k & 重复项计数分隔符 & cnt(k)
            d.add k, ""
        End If
    Next k
End Function
'https://brettdotnet.wordpress.com/2012/03/30/convert-a-collection-to-an-array-vba/
Function CollectionToArray(c As Collection) As Variant()
    Dim A() As Variant: ReDim A(1 To c.Count)
    Dim i As Integer
    For i = 1 To c.Count
        A(i) = c.Item(i)
    Next
    CollectionToArray = A
End Function


Private Function get自动填充的范围(ByVal rng As Range) As Range
    Dim lastCol As Long: lastCol = rng.Column + rng.Columns.Count - 1 'rng最后一列列号
    Dim regionRng As Range: Set regionRng = rng.CurrentRegion
    Dim lastRow As Long: lastRow = regionRng.Row + regionRng.Rows.Count - 1 '与lastCol同理
    Dim c As Range: Set c = rng.Parent.Cells(lastRow, lastCol)   '用于定位的单元格
    Set get自动填充的范围 = Range(rng, c)
End Function
Private Sub TEST_get自动填充的范围()
    Dim rng As Range, rng2 As Range
    Set rng = Range("B2:C2")
    rng.AutoFill get自动填充的范围(rng), xlFillSeries
End Sub
Sub myAutoFill(rng As Range, Optional myType As XlAutoFillType = xlFillDefault, Optional convert2value As Boolean)
    Dim rng2 As Range: Set rng2 = get自动填充的范围(rng)
    rng.AutoFill rng2, myType
    If convert2value Then
        rng2.Copy
        rng2.PasteSpecial xlPasteValues
    End If
End Sub
Private Sub TEST_myAutoFill()
    myAutoFill Range("B2:C3"), xlFillSeries
End Sub
