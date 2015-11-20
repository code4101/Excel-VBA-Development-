开发背景
====
资产稽核中，需要对表格模板相同的大约5人的表格，进行数据汇总。
由于“表头”相同，关键是对内容的汇总，所以计划开发对相同表格（或者说前几行相同的表格），能将下方数据清除、汇至一起等功能。

使用方法
```Visual Basic
Private Sub TEST_CdataSheet()
    Dim ob As New CdataSheet
    ob.SetUp ActiveSheet, 2

    ob.ClearData

    Dim i As Long
    For i = 1 To 6
        ob.AddData Sheets(CStr(i)), xlPasteValues
    Next i
End Sub

Private Sub ClearData()
    Dim ob As New CdataSheet
'    ob.SetUp Sheets("ALL"), 2
'    ob.ClearData
    ob.SetUp Sheets("SGY"), 2
    ob.ClearData
    ob.SetUp Sheets("ZPL"), 2
    ob.ClearData
    ob.SetUp Sheets("XLN"), 2
    ob.ClearData
    ob.SetUp Sheets("HLM"), 2
    ob.ClearData
End Sub
```
