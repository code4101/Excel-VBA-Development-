Attribute VB_Name = "stdvba"
'code4101 ���¸�����:2015��11��20�� 21:00+

Enum ��ɫ��
    ��׼�ֶ���ɫ = 15773696 'RGB(0, 176, 240)   ��ɫ
    �����ֶ���ɫ = 5296274  'RGB(146, 208, 80)  ��ɫ
    �ֶη�����ɫ = 65535    'RGB(255, 255, 0)   ��ɫ
End Enum


Function ����ʹ�õ����һ����Ԫ��(ByVal x As Range) As Range
''�㷨˼·: ʹ������ָ�룬yָ��x����һ����Ծ�㣬zָ��y����һ����Ծ��
'' ��y��z��ָ��ͬʱ����ʱx��Ϊ���ݽ�β
'    Dim y As Object, z As Object
'    Set y = x.End(xlDown)
'    Set z = y.End(xlDown)
'    Do While y.Address <> z.Address
'        Set x = y
'        Set y = z
'        Set z = z.End(xlDown)
'    Loop
'    Set ����ʹ�õ����һ����Ԫ�� = x
'2015/7/29�ո���
    '���ȸ���x�Ĳ�������objectΪrange
    'Ȼ��ı��㷨
    Dim st As Worksheet: Set st = x.Parent
    With st
        Set ����ʹ�õ����һ����Ԫ�� = .Cells(.Cells(.Rows.Count, x.Column).End(xlUp).Row, x.Column)
    End With
End Function

'��x��xĩβ�����е����ݿ�����y��Ԫ�񼰺���
Sub �������ܿ���(x As Range, y As Range)
    Range(x, ����ʹ�õ����һ����Ԫ��(x)).Copy y
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''sheet�������Ϣ��ȡ''''''''''''''''''''''''''''''''''''''''''''
' ��(i,j)��ʼ�����±����������մ������һ��
Function ���һ��(һ�ű�, Optional ByVal i As Integer = 1, Optional ByVal j As Integer = 1)
    Do While һ�ű�.Cells(i, j) <> ""
        i = i + 1
    Loop
    ���һ�� = i - 1
End Function


' ��(i,j)��ʼ�����ұ����������մ������һ��
Public Function ���һ��(һ�ű�, Optional ByVal i As Integer = 1, Optional ByVal j As Integer = 1)
    Do While һ�ű�.Cells(i, j) <> ""
        j = j + 1
    Loop
    ���һ�� = j - 1
End Function

' ��(x1,y1)��(x2,y2)����ָ������
' ���ҵ����򽫽���洢��x1,y1
' ���ز���ֵ:true����ɹ�,false����ʧ��
Function ��һ����Χ�ڲ���ָ���ı�����λ��(һ�ű�, ����ֵ, ByRef x1, ByRef y1, x2 As Integer, y2 As Integer) As Boolean
    For i = x1 To x2
        For j = y1 To y2
            If һ�ű�.Cells(i, j) = ����ֵ Then
                x1 = i
                y1 = j
                ��һ����Χ�ڲ���ָ���ı�����λ�� = True
                Exit Function
            End If
        Next j
    Next i
    ��һ����Χ�ڲ���ָ���ı�����λ�� = False
End Function

'ʹ�þ���:  Debug.Print is�ñ����(Workbooks("������ȡ��.xlsb"), "Sheet1")
Function is�ñ����(������, ���� As String) As Boolean
    is�ñ���� = False
    For i = 1 To ������.Sheets.Count
        If ���� = ������.Sheets(i).name Then
            is�ñ���� = True
            Exit Function
        End If
    Next i
End Function



' ɾ����i�У���j����ĵ�Ԫ��
Sub ��������(i As Integer, j As String)
    Rows(i & ":1048576").Delete Shift:=xlUp
    Columns(j & ":XFD").Delete Shift:=xlToLeft
End Sub

' ���յ�һ�е�������Զ������(x2,y2)
Sub �Զ����(ByVal x1 As Integer, ByVal y1, ByVal x2 As Integer, ByVal y2)
    Range(Cells(x1, y1), Cells(x1, y2)).Select
    Selection.AutoFill Destination:=Range(Cells(x1, y1), Cells(x2, y2)), Type:=xlFillDefault
End Sub
Sub �Զ����2(rng As Range)
    Dim rng2 As Range   'Ҫ����ȫ����Χ
    Set rng2 = rng.Parent.UsedRange '�Ȼ������ʹ�÷�Χ
    
    rng.AutoFill Destination:=Range("A3:A208")
End Sub

Function ת��ֵ(A) As Double
    If A = "" Then
        ת��ֵ = 0#
    Else
        ת��ֵ = A
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''��ѧ��(math.h)'''''''''''''''''''''''''''''''''''''''''''''
Function pow(A, m)
' a�Ƿ���,m�ǲ�С��1����������
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


Function ����(ByVal �к� As Long) As String
    If �к� = 0 Then
        ���� = "#N/A"
        Exit Function
    End If
    Dim s As String: s = addr(Cells(1, �к�))
    ���� = Left(s, Len(s) - 1)
End Function

'' ���е����ֱ��ת��Ϊ��ĸ���
'Function ����(�к� As Integer) As String
'    Do While �к� > 25
'        �к� = �к� - 26
'        ���� = ���� + "Z"
'    Loop
'    If �к� > 0 Then ���� = ���� + Chr(64 + �к�)
'End Function

'''''''''''''''''''''''''''''''''''''''''''''�㷨��(algorithm)'''''''''''''''''''''''''''''''''''''''''''''
' ��Դ:http://stackoverflow.com/questions/152319/vba-array-sort-function
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

Function CZ(����ֵ As String, ����ֵ�������� As Range, Optional Ŀ��ֵ������ As Variant, Optional ȷ�Ϸ��صڼ���Ŀ��ֵ As Integer = 1, Optional ģ������ As Integer = 1) As String
    Application.Volatile
    Dim i As Long, r As Range, R1 As Range, Str As String, L As Long
    Dim CZFS As Long
    Dim st As String, p As Long
    
    If ģ������ = 2 Then   '1������ģ�����ң�2������ģ������
       st = ""
       For p = 1 To Len(����ֵ)
           st = st & Mid(����ֵ, p, 1) & "*"
       Next p
       ����ֵ = Left(st, Len(st) - 1)
    End If
    
    If ģ������ > 0 Then CZFS = xlPart Else CZFS = xlWhole
    
    Dim sh As Worksheet, SH1 As Worksheet
    
      
    With ����ֵ��������(1).Resize(����ֵ��������.Rows.Count, 1)
    If .Cells(1) = ����ֵ Then Set r = .Cells(1) Else Set r = .Find(����ֵ, LookIn:=xlValues, lookat:=CZFS)
     If Not r Is Nothing Then
        Set sh = r.Parent
     
     
        If TypeName(Ŀ��ֵ������) = "Range" Then
           Set R1 = Ŀ��ֵ������
           Set SH1 = R1.Parent
           L = Ŀ��ֵ������.Column
        Else
           L = Ŀ��ֵ������
           If L = 0 Then L = r.Column
        End If
     
        Str = r.Address
        Do
            i = i + 1
            If i = ȷ�Ϸ��صڼ���Ŀ��ֵ Then
              If Not SH1 Is Nothing Then CZ = SH1.Cells(r.Row, L) Else CZ = Cells(r.Row, L)
              Exit Function
            End If
            Set r = ����ֵ��������.Find(����ֵ, r, lookat:=CZFS)
        Loop While Not r Is Nothing And r.Address <> Str
    End If
End With
End Function


'''''''''''''''''''''''''''''''''''''''''''''�ַ�����(algorithm)'''''''''''''''''''''''''''''''''''''''''''''
Function �ַ������ƶ�(ȫ�� As String, ��� As String)
    ����1 = Len(ȫ��)
    ����2 = Len(���)
    
    �ַ������ƶ� = 0
    k = 1
    For i = 1 To ����2
            
        ' �� k~����1 �� mid(���,i,1), ��λ��Ϊj
        pos = -1
        For j = k To ����1
            If Mid(���, i, 1) = Mid(ȫ��, j, 1) Then
                pos = j
                Exit For
            End If
        Next j
        
        If pos <> -1 Then
            �ַ������ƶ� = �ַ������ƶ� + 1
            k = pos
        End If
        
    Next i
    
End Function


Function �ַ�����(rng As Range, Optional �зָ��� As String = ",", Optional �зָ��� As String = ";") As String

    For i = 1 To rng.Rows.Count
        
        If i <> 1 Then �ַ����� = �ַ����� & �зָ���
        �ַ����� = �ַ����� & rng.Cells(i, 1)
        
        For j = 2 To rng.Columns.Count
            �ַ����� = �ַ����� & �зָ��� & rng.Cells(i, j)
        Next j
        
    Next i
    
End Function


' �����x��y����һ����Ԫ��
' maxn��ѡ������ʾ���ӳ��ļ�ֵ��
Function һ�Զ��ֵ����(x, y, Optional maxn As Integer = 20, Optional �ָ��� As String = ",")
    '(0)�����ֵΪ��,�򷵻ؿ�
    If x = "" Then
        һ�Զ��ֵ���� = ""
        Exit Function
    End If
    '(1) ����������
    Dim i
    For i = 1 To maxn
        If (x.Offset(i, 0) <> "" And x.Offset(i, 0) <> x) Or y.Offset(i, 0) = "" Then
            Exit For
        End If
    Next i
    '��2�������Ӻ����Է�Χ�ڵ�ֵ����ƴ��
    һ�Զ��ֵ���� = �ַ�����(Range(y, y.Offset(i - 1, 0)), "", �ָ���)
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

'���������(0~9)����
Function CleanString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
     .Global = True
     .Pattern = "[^\d]+"
    CleanString = .Replace(strIn, vbNullString)
    End With
End Function


'���������2015��07��30��
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


'�ú���֧��name��partName�÷ֺŸ������������ȼ������ֶ��������Ķ��ֶβ�ѯ
Function findcel(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Range
'(1)����name���Բ���Ϊ��
    If name = "" Then Exit Function

    Dim arr1, arr2
'(2)partName����Ϊ�գ���Ϊ�˺�������ͳһ������Ҫ��Ԥ������
    arr1 = Split(partName, ";")
    If isEmptyArr(arr1) Then
        ReDim arr1(1 To 1)
        arr1(1) = ""
    End If
    
'(3)��ʼѭ������,ֻҪ�ҵ���һ������⼴��
    arr2 = Split(name, ";")
    For Each A1 In arr1
        For Each A2 In arr2
            Set findcel = findcel_base(st, A2, A1)
            If Not (findcel Is Nothing) Then Exit Function
        Next A2
    Next A1
End Function

Function findcel_base(ByVal st As Worksheet, ByVal name As String, Optional ByVal partName As String) As Range
    Dim rng As Range '���ҵķ�Χ
    Set rng = st.UsedRange
    
    'Debug.Print "findcel_base�����������ڹ�����", st.Parent.name
    Dim rng2 As Range, t As Range
'(1)�ȶ�λ�߼���ͷ���з�Χ
    If partName <> "" Then
        Set t = rng.Find(partName, lookat:=xlPart)
        '�����һ���Ǻϲ���Ԫ����ʱ������Ҳ�����bug
        If rng.Cells(1, 1) = partName Then Set t = rng.Cells(1, 1)
        '���ȷʵ�Ҳ������˳�����
        If t Is Nothing Then Exit Function
        
        '��������ҵ��ˣ�������ҵ���(�ϲ�)��Ԫ��������
        Set rng2 = st.Range(rng.Cells(1, t.Column), rng.Cells(st.Rows.Count, t.Offset(0, 1).Column - 1))
        Set rng = Intersect(rng, rng2)  'Range�Ľ�
    End If

'(2)Ȼ��Ϳ���ֱ����rng������ͷ����
    Set t = rng.Find(name, lookat:=xlWhole)                        '�ܵ�Ԫ��ƥ���ҵ������յ�Ԫ����
    If t Is Nothing Then Set t = rng.Find(name, lookat:=xlPart)    '������в��ֲ���
    If name = rng.Cells(1, 1) Then Set t = rng.Cells(1, 1)          '�����һ����Ԫ��ֵ���㣬ǿ�ƽ���������Ϊcells(1,1)
    '���ǰ�漸�������û�ҵ�����cells(1,1)�Ƿ�����ģ��ƥ��
    If t Is Nothing And rng.Cells(1, 1) Like "*" & name & "*" Then Set t = rng.Cells(1, 1)
    
    'If Not (t Is Nothing) Then Debug.Print name & "��" & t.Address
    Set findcel_base = t
End Function

Private Function isEmptyArr(arr) As Boolean  '
    isEmptyArr = True
    For Each A In arr
        isEmptyArr = False
        Exit For
    Next A
End Function


'��õ�Ԫ���λ��(ȥ���������÷�)
Function addr(cell) As String
    addr = Replace(cell.Address, "$", vbNullString)
End Function


'oriStr��ԭʼ�ļ�����·����optPath�ǲο�·��
'���oriStr�Ѿ�����Ч���ļ������򷵻�ԭֵ
'���򣬽�oriStr������Ŀ¼���õ�optPath
'�������Ч·�����᷵�ؿ��ַ���
Function getfn(ByVal oriStr As String, Optional optPath As String) As String
'�ο�����:http://zhidao.baidu.com/link?url=9qQA8dJddTAGsmuPyrKpl6IQbBnxI7PNY9-os-WZhjsj2k5V4-d95nfR6GFlr8hL3uW-RCrL_St1EouTmJiX7bU5m6KQZDBQU0_VGY_31EW
    Dim fso As Object
    Dim res As String
    
    '���fsoò�ƺ����ܣ�����������Ѵ򿪣�Ҳ����ȷʶ���Ѵ���
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
Private Sub hasStn����()
    'Debug.Print hasStn(ActiveWorkbook, "2-2�̶��ʲ�������")
    Debug.Print is�ñ����(ActiveWorkbook, "2-2�̶��ʲ�������")
End Sub

'�Զ�����������
Sub autoFit(st As Worksheet)
    With st.Cells
        .EntireColumn.autoFit
        .EntireRow.autoFit
    End With
End Sub


'is���� Ϊ True:�����ʼ��
'is���� Ϊ False:����ʼ��
Sub �����ʼ���뷴��ʼ��(is���� As Boolean)
    Static tt As Double
    If is���� Then
        tt = Timer  '��ʱ��
        With Application
            .ScreenUpdating = False     '�ر���Ļ���¼ӿ�ִ���ٶ�
            .DisplayAlerts = False
        End With
    Else
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
        End With
        MsgBox "����������ϣ���ʱ " & Format(Timer - tt, "#0.00") & "��", Title:="��ʱ��"
    End If
End Sub


'Function getWb(ByVal wbn) As Workbook
'    Dim st0 As Worksheet: Set st0 = ActiveSheet       '���ڴ洢ԭ�������
'
'    Debug.Print Chr(10); "getWb�����У����Դ��ļ���"; CStr(wbn)
'    wbn = getfn(CStr(wbn))
'
'    If wbn = "" Then
'        Debug.Print "�����ļ�·����Ч"
'        Exit Function
'    Else
'        Debug.Print "���ļ�����Ч"
'        '������ļ��Ѵ�
'        For Each wb In Workbooks
'            If wbn Like ("*" & wb.name & "*") Then
'                Debug.Print "�������ļ��Ѿ����ڴ�״̬"
'                Set getWb = wb
'                Exit Function
'            End If
'        Next wb
'
'        '���δ��
'        Set getWb = Workbooks.Open(wbn, False, False)    'ʹ����Ĭ�ϵ�ϰ���Բ������ļ�
'        If getWb Is Nothing Then
'            Debug.Print "������Ȼ��ʧ��";
'        Else
'            st0.Activate                                '����ԭ"�������"
'            Debug.Print "�򿪳ɹ�:"; getWb.name
'        End If
'    End If
'
'End Function

Function getWb(ByVal wbn) As Workbook
    Dim wb As Workbook

'(1)ͳһΪ����·��
    wbn = CStr(wbn)
    If Mid(wbn, 2, 1) <> ":" Then wbn = ActiveWorkbook.path & "\" & wbn

'(2)���Ѿ��򿪵��ļ��Ա�
    Dim wb1 As Workbook
    Dim path As String, name As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.GetParentFolderName(wbn)
    name = fso.GetFileName(wbn)
    
    For Each wb1 In Workbooks
        If wb1.path = path And wb1.name = name Then
            Debug.Print "���ļ��Ѿ����ڴ�״̬:", wbn
            Set getWb = wb1
            Exit Function
        ElseIf wb1.name = name And wb1.path <> path Then
            Debug.Print "�����������ظ�,�޷��򿪸��ļ�:", wbn
            Exit Function
        End If
    Next wb1

'(3)���Դ�
    On Error Resume Next
    Set wb = Workbooks.Open(wbn, False)
    
    If wb Is Nothing Then
        Debug.Print "�ļ�����Ч,���ļ�δ��:", wbn
    Else
        Debug.Print "�ɹ����ļ�:", wbn
        Set getWb = wb
    End If
End Function

Private Sub ����getWb()
    Dim wb As Workbook
    Set wb = getWb("�й��ƶ�31����02������ϸ��0803-����.xlsx")
    
    Debug.Print "��ǰ����Ĺ�����·��:", ActiveWorkbook.path
    If wb Is Nothing Then
        Debug.Print "�ۺϲ��Խ��:δ���ļ�"
    Else
        Debug.Print "�ۺϲ��Խ��:���ļ�", wb.path, wb.name
    End If
End Sub

Private Sub fso�������()
    s = ActiveWorkbook.path & "\aaa"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Debug.Print fso.GetParentFolderName(s)
    Debug.Print fso.GetFileName(s)
End Sub

Function has����(s As String) As Boolean
    has���� = False
    Dim t As Long, i As Long
    
    For i = 1 To Len(s)
        t = Asc(Mid(s, i, 1))
        If t < 1 Or t > 127 Then
            has���� = True
            Exit Function
        End If
    Next i
End Function


Function get����Text( _
    ByVal c As Collection, _
    Optional ���� As Boolean = True, _
    Optional �ظ������ As Boolean = True, _
    Optional ��ָ��� As String = ", ", _
    Optional �ظ�������ָ��� = "") As String
'(1)
    If �ظ������ Then
        Set c = ��Collection�ظ������(c, �ظ�������ָ���)
    End If
'(2)
    If ���� Then
        Set c = ��Collection����(c)
    End If
'(3)
    Dim ans As String
    For Each cc In c
        If ans = "" Then
            ans = cc
        Else
            ans = ans & ��ָ��� & cc
        End If
    Next cc
'(4)return
    get����Text = ans
End Function
Function ��Collection����(c As Collection) As Collection
'ʵ��ԭ���Ƚ�����ת������Array����Arrayʹ��quicksort��Ȼ���ٽ��ź����Array����µ�Colllection
    Set ��Collection���� = New Collection
    Dim A() As Variant
    A = CollectionToArray(c)
    Call QuickSort(A, LBound(A), UBound(A))
    For Each aa In A
        ��Collection����.add aa
    Next aa
End Function
Function ��Collection�ظ������(c As Collection, Optional �ظ�������ָ��� = "") As Collection
'(1)�ȼ����ÿһ�������
    Dim cnt As Object: Set cnt = CreateObject("Scripting.Dictionary") '���ڽ����ֵ丨��
    For Each k In c
        cnt(k) = cnt(k) + 1
    Next k
'(2)����µļ���
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Set ��Collection�ظ������ = New Collection
    For Each k In c
        If Not d.Exists(k) Then 'd�洢�Ѿ�visited����
            ��Collection�ظ������.add k & �ظ�������ָ��� & cnt(k)
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


Private Function get�Զ����ķ�Χ(ByVal rng As Range) As Range
    Dim lastCol As Long: lastCol = rng.Column + rng.Columns.Count - 1 'rng���һ���к�
    Dim regionRng As Range: Set regionRng = rng.CurrentRegion
    Dim lastRow As Long: lastRow = regionRng.Row + regionRng.Rows.Count - 1 '��lastColͬ��
    Dim c As Range: Set c = rng.Parent.Cells(lastRow, lastCol)   '���ڶ�λ�ĵ�Ԫ��
    Set get�Զ����ķ�Χ = Range(rng, c)
End Function
Private Sub TEST_get�Զ����ķ�Χ()
    Dim rng As Range, rng2 As Range
    Set rng = Range("B2:C2")
    rng.AutoFill get�Զ����ķ�Χ(rng), xlFillSeries
End Sub
Sub myAutoFill(rng As Range, Optional myType As XlAutoFillType = xlFillDefault, Optional convert2value As Boolean)
    Dim rng2 As Range: Set rng2 = get�Զ����ķ�Χ(rng)
    rng.AutoFill rng2, myType
    If convert2value Then
        rng2.Copy
        rng2.PasteSpecial xlPasteValues
    End If
End Sub
Private Sub TEST_myAutoFill()
    myAutoFill Range("B2:C3"), xlFillSeries
End Sub


