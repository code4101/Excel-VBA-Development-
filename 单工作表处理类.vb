'*** 类介绍 ***
'对表格进行规范化处理的通用性函数
'主要是找出某一字段名，与该字段数值的匹配模式，来规范化处理


'*** 该类使用方法举例 ***
'Private Sub 类测试()
'    Set ob = New 单工作表处理类
'
'    '通用方法:一般情况下是这样调用的
'    'Set wb = Workbooks.Open("C:\Users\chen\Downloads\aaa.xlsx")
'    'Call ob.初始化设置(wb.Sheets("5-3租赁铁塔"), "物理站址编号", "*福建*")
'
'    '测试方法:只处理当前激活表
'    Call ob.初始化设置(ActiveSheet, "物理站址编号", "*福建*") '*** 这里可以根据不同问题设置特定的匹配模式 ***
'
'    Call ob.表格标准化处理                                    '*** 调用该方法能对表格进行处理 ***
'End Sub


Option Explicit
Option Base 1

Dim thisSheet As Worksheet
Dim 关键字段名 As String
Dim 有效模式 As String


Dim 关键字段所在列 As Long


' 默认表格规则：
'以当前激活工作表为对象
'以当前激活单元格内容为字段名，不空为有效模式
Private Sub Class_Initialize()
    Set thisSheet = ActiveSheet
    关键字段名 = ActiveCell
    有效模式 = "?"
End Sub

Sub 初始化设置(输入_thisSheet As Worksheet, 输入_关键字段名 As String, 输入_有效模式 As String)
    Set thisSheet = 输入_thisSheet
    关键字段名 = 输入_关键字段名
    有效模式 = 输入_有效模式
End Sub

Private Function is该表存在(表名 As String) As Boolean
    is该表存在 = False
    Dim i As Long
    For i = 1 To ActiveWorkbook.Sheets.Count
        If 表名 = ActiveWorkbook.Sheets(i).name Then
            is该表存在 = True
            Exit Function
        End If
    Next i
End Function

Private Sub 如果不存在该表则在末尾新建(表名 As String)
    If is该表存在(表名) = False Then
        Worksheets.Add after:=Sheets(Sheets.Count)
        ActiveSheet.name = 表名
        'Debug.Print Now, "创建如下工作表:", 表名
    End If
End Sub

Private Sub 把内容加入某表(表名 As String, 标签行数组 As Variant, rng As Range)
    'If rng.Rows.Count < 2 Then Exit Sub ' 如果内容只有一行则退出sub（这一行是必有的表头~）
    
    Call 如果不存在该表则在末尾新建(表名)
    Dim 无效行_表格 As Worksheet
    Set 无效行_表格 = Sheets(表名)


    Dim 填写行 As Long
    With 无效行_表格
        填写行 = .UsedRange.Rows.Count
        If 填写行 > 1 Then ' 如果填写行不在第一行，先移动三行
            .Range(.Cells(填写行 + 1, 1), .Cells(填写行 + 3, 1)) = ""
            填写行 = 填写行 + 3
        End If
        .Range(.Cells(填写行, 1), .Cells(填写行, UBound(标签行数组))) = 标签行数组
        rng.Copy .Rows(填写行 + 1)
    End With
    'Debug.Print Now, "执行了sub:把内容加入某表"
End Sub



Sub 表格标准化处理()
'过程功能
'1、程序首先会找到关键字段所在行列，如果为合并单元格。其所在行范围取消合并单元格，所有字段名填充在下方。
'2、关键字段所在及以上行，和以下所有无效行，复制到“无效行”表；去除关键字段所在行后，剩余行删除

    Dim i As Long, j As Long
'(1)表头处理
    Dim 要删除的rng As Range, 要写出去的rng As Range
    
    Dim c As Object
    Dim c的下面所在行 As Integer
    
    With thisSheet
        Set c = .UsedRange.Find(关键字段名)
        c的下面所在行 = c.Offset(1, 0).Row
        关键字段所在列 = c.Column
        
        '如果是合并表头，则标准化处理
        If c.MergeCells Then
            '在原表头下面插入一行，制作一个新表头
            c.Offset(1, 0).EntireRow.Insert Shift:=xlDown
            
            Debug.Print Now, "对 " & thisSheet.name & " 规范处理", "表格原UsedRange行数:" & .UsedRange.Rows.Count
            
            For j = 1 To .UsedRange.Columns.Count
                .Cells(c的下面所在行, j) = .Cells(c的下面所在行, j).End(xlUp)
            Next j
            Set c = c.Offset(1, 0)
            c的下面所在行 = c.Offset(1, 0).Row
        End If
        
        ' 如果表头不在第一行，那么把第一行前面的都删除
        If c.Row > 1 Then
            Set 要删除的rng = .Rows(1 & ":" & c.Row - 1)
        End If
        
    '(2)无效行处理
        For i = c的下面所在行 To .UsedRange.Rows.Count
            If Not .Cells(i, 关键字段所在列) Like 有效模式 Then '如果这格没有福建字眼，需要剪切走
                If 要删除的rng Is Nothing Then
                    Set 要删除的rng = .Rows(i)
                Else
                    Set 要删除的rng = Union(要删除的rng, .Rows(i))
                End If
            End If
        Next i
        
    '(3)先复制内容，再删除
        If 要删除的rng Is Nothing Then
            Exit Sub    '如果没有需要删除的内容，则直接结束sub
        End If
    
        Set 要写出去的rng = Union(要删除的rng, .Rows(c.Row))
        Call 把内容加入某表("无效行", Array("工作表名:", .name, "", "处理时间:", Now), 要写出去的rng)
        要删除的rng.Delete
    End With

'(4)其它
    '(4.1)清除分级显示
    thisSheet.Cells.ClearOutline

    '(4.2)冻结首行
    Dim 屏幕更新原状态 As Boolean
    屏幕更新原状态 = Application.ScreenUpdating
    
    Application.ScreenUpdating = True   '无论原状态如何，一律开启
    thisSheet.Activate
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    Cells(1, 1).Activate
    Application.ScreenUpdating = 屏幕更新原状态 '恢复屏幕更新原状态
    
    '(4.x)还有行列autofit的可以调整，不过暂时维持表格原状吧
'    With Cells
'        .EntireColumn.AutoFit
'        .EntireRow.AutoFit
'    End With
    
    '(4.3)无效行表格的行高自动调整（做铁塔32表时发现有时行高会极不正常）
    Sheets("无效行").Cells.EntireRow.AutoFit
End Sub

Sub 字段改名(旧名 As String, 新名 As String)
    thisSheet.Rows(1).Replace 旧名, 新名
End Sub


'删除字段功能，本来一开始的设计是：输入的参数是一个数组，进入函数后，转化为字典，进行批量删除
'不过仔细想一下，觉得这样效率未必更好，每次都要转字典也是要代价的。
'总支，目前还是先用简单的方法实现基本功能。
' 输入一个数组，在该表中，遍历所有字段，如果该字段在arr中出现则删除
Sub 删除字段(ByVal 字段名 As String)
    Dim p As Object
    Set p = thisSheet.Rows(1).Find(字段名)
    If Not p Is Nothing Then
        thisSheet.Columns(p.Column).Delete
    End If
End Sub

