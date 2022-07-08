Attribute VB_Name = "Logic"



Sub sum_array(counter() As Integer, n As Integer, ByRef sum As Integer)
    Dim i As Integer
    i = 0
    sum = 0
    While (i < n)
        sum = sum + counter(i)
        i = i + 1
        Wend
        sum = sum - 1
End Sub



Sub four_gen(ByRef para As pairs, ByRef chet As Test, p As Integer, k As Integer, ByRef bool1() As Boolean)
Dim ws As Worksheet, ws1 As Worksheet
Dim n As Integer
ReDim Preserve chet.quest_name(k - 1)

    Set ws = ThisWorkbook.Sheets("Четверки")
    Set ws1 = ThisWorkbook.Sheets("Слова и группы")
    Dim used() As Boolean
    Dim counter() As Integer
    Dim irow As Integer, _
        icol As Integer
        irow = 2
        st_diapazon = 0
        i = 1
        While ws1.Cells(1, i) <> ""
            While ws1.Cells(irow, i) <> ""
                ReDim Preserve counter(i - 1)
                irow = irow + 1
                counter(i - 1) = irow - 2
                Wend
            irow = 2
            i = i + 1
        Wend
        i = 1
        ws.Cells(5, 10).Value = "=RANDBETWEEN(0," & CStr(para.count - 1) & ")"
        n = ws.Cells(5, 10).Value
        ws.Cells(5, 10).Clear
        While (bool1(n) = True)
        ws.Cells(5, 10).Value = "=RANDBETWEEN(0," & CStr(para.count - 1) & ")"
        n = ws.Cells(5, 10).Value
        ws.Cells(5, 10).Clear
        Wend
        bool1(n) = True
        ReDim Preserve used(para.count - 1)
 If (p = 1) Then
        chet.quest_name(k - 1).right = para.item(n).word
        chet.quest_name(k - 1).question = para.item(n).translation
        used(n) = True
        Else
        chet.quest_name(k - 1).right = para.item(n).translation
        chet.quest_name(k - 1).question = para.item(n).word
        used(n) = True
        End If
        i = i + 1
        Dim r As Integer
        r = n
        
Dim l As Integer, j As Integer

 While (i < 5)
    If (p = 1) Then
        j = 0
        l = 0
        If (r < counter(0)) Then
            st_diapazon = 0
            l = counter(0) - 1
        Else
            
        While (l < r)
            Call sum_array(counter, j, l)
            j = j + 1
            Wend
            j = j - 1
            st_diapazon = l - counter(j - 1) + 1
        End If
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            While (used(n) = True Or StrComp(para.item(n).word, "", 1) = 0)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            chet.quest_name(k - 1).wrong(i - 2) = para.item(n).word
            used(n) = True
        End If
                
    If (p = 2) Then
        j = 0
        l = 0
        If (r < counter(0)) Then
            st_diapazon = 0
            l = counter(0) - 1
        Else
            
        While (l < r)
            Call sum_array(counter, j, l)
            j = j + 1
            Wend
            j = j - 1
            st_diapazon = l - counter(j - 1) + 1
        End If
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            While (used(n) = True Or StrComp(para.item(n).word, "", 1) = 0)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            chet.quest_name(k - 1).wrong(i - 2) = para.item(n).word
            used(n) = True
            End If
            
 i = i + 1
 Wend
 chet.num = chet.num + 1
End Sub


Sub four_gen1(ByRef para As pairs, ByRef chet As Test, p As Integer, d As Integer, k As Integer, ByRef bool1() As Boolean)
Dim ws As Worksheet, ws1 As Worksheet
Dim n As Integer

    ReDim Preserve chet.quest_name(k - 1)
    Set ws = ThisWorkbook.Sheets("Четверки")
    Set ws1 = ThisWorkbook.Sheets("Слова и группы")
    Dim used() As Boolean
    Dim counter() As Integer
    Dim irow As Integer, _
        icol As Integer
        irow = 2
        st_diapazon = 0
        i = 1
        While ws1.Cells(1, i) <> ""
            While ws1.Cells(irow, i) <> ""
                ReDim Preserve counter(i - 1)
                irow = irow + 1
                counter(i - 1) = irow - 2
                Wend
            irow = 2
            i = i + 1
        Wend
        i = 1
        ReDim Preserve used(para.count - 1)
        Dim l As Integer
        Call sum_array(counter, d, l)
        st_diapazon = l - counter(d - 1) + 1
        ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
        n = ws.Cells(5, 10).Value
        ws.Cells(5, 10).Clear
        While (bool1(n) = True)
        ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
        n = ws.Cells(5, 10).Value
        Wend
        bool1(n) = True
        ws.Cells(5, 10).Clear
        ReDim Preserve used(para.count - 1)
 If (p = 1) Then
        'ws.Cells(1, i) = para.item(n).word
        chet.quest_name(k - 1).right = para.item(n).word
        chet.quest_name(k - 1).question = para.item(n).translation
        used(n) = True
        
        Else
        'ws.Cells(1, i) = para.item(n).translation
        chet.quest_name(k - 1).right = para.item(n).translation
        chet.quest_name(k - 1).question = para.item(n).word
        used(n) = True
        End If
        i = i + 1
        

        
 While (i < 5)
    If (p = 1) Then
            While (used(n) = True Or StrComp(para.item(n).word, "", 1) = 0)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            chet.quest_name(k - 1).wrong(i - 2) = para.item(n).word
            used(n) = True
        End If
                
    If (p = 2) Then
            While (used(n) = True Or StrComp(para.item(n).word, "", 1) = 0)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(st_diapazon) & "," & CStr(l) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            chet.quest_name(k - 1).wrong(i - 2) = para.item(n).word
            used(n) = True
            End If

            
 i = i + 1
 Wend
 chet.num = chet.num + 1
End Sub
    
Sub all_four(ByRef para As pairs, ByRef fo As Test, language As Integer)
Dim m As Integer
m = para.count
Dim global_bool() As Boolean
ReDim Preserve global_bool(m)
Dim ws1 As Worksheet
Set ws1 = ThisWorkbook.Sheets("Слова и группы")
Dim n As Integer
n = 1
While ws1.Cells(1, n) <> ""
    n = n + 1
    Wend
m = 1
While (m < n)
    Call four_gen1(para, fo, language, m, m, global_bool)
    m = m + 1
    Wend
m = n
While (m <= para.count)
    Call four_gen(para, fo, language, m, global_bool)
    m = m + 1
    Wend
End Sub



