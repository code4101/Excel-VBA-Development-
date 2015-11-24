Attribute VB_Name = "t1029OpenAndCloseWksEffec"
Sub main()
    Call 程序初始化与反初始化(True)

    Dim fn As String
    Dim wb As Workbook
    Dim files As String: files = "W:\中国铁塔\10月 稽核_批量读取单站数据\data\1\"
    
    fn = Dir(files & "*.xls")
    Do While fn <> ""
        Set wb = Workbooks.Open(files & fn, False, True)    '唉，用workbooks.open并没有比getobject快~~
        If Not (wb Is Nothing) Then wb.Close
        fn = Dir()
    Loop

    Call 程序初始化与反初始化(False)
End Sub

