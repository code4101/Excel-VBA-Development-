[Excel VBA 如何快速学习？](http://www.zhihu.com/question/20870802/answer/54998361)
----
比如，你在Excel里面使用了VBA，那么Excel就是宿主。这个概念也适用于VBScript，对应的宿主可以认为是Windows本身（实际上是Windows的一个组件：Windows脚本宿主WSH）。
 
http://www.csidata.com/custserv/onlinehelp/vbsdocs/VBSTUTOR.HTM
https://msdn.microsoft.com/en-us/library/d1wf56tt(v=vs.84).aspx

```vb
Option Explicit

Dim app, wb, st
Dim row, col

Set app = WScript.CreateObject("Excel.Application")
app.Visible = True
Set wb = app.Workbooks.Add

Set st = wb.Worksheets(1)

For row = 1 To 10
    For col = 1 To 10
        st.Cells(row, col).Value = CInt(Int((100 * Rnd()) + 1))
    Next
Next
Set st = wb.Worksheets(2)
st.Range("A1:J10").Formula = "=int(rand()*100+1)"
```
