Attribute VB_Name = "t1029OpenAndCloseWksEffec"
Sub main()
    Call �����ʼ���뷴��ʼ��(True)

    Dim fn As String
    Dim wb As Workbook
    Dim files As String: files = "W:\�й�����\10�� ����_������ȡ��վ����\data\1\"
    
    fn = Dir(files & "*.xls")
    Do While fn <> ""
        Set wb = Workbooks.Open(files & fn, False, True)    '������workbooks.open��û�б�getobject��~~
        If Not (wb Is Nothing) Then wb.Close
        fn = Dir()
    Loop

    Call �����ʼ���뷴��ʼ��(False)
End Sub

