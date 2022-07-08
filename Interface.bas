Attribute VB_Name = "Interface"
Sub xls_read(ByRef para As pairs)
    Dim b As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Слова и группы")
    Dim irow As Integer, _
        icol As Integer
        irow = 2
        para.count = 0
        i = 1
        While ws.Cells(1, i) <> ""
            While ws.Cells(irow, i) <> ""
                ReDim Preserve para.item(para.count)
                Dim a() As String
                a = Split(ws.Cells(irow, i), "-")
                For j = 0 To para.count - 1
                    If (StrComp(a(0), para.item(j).word, 1) = 0 Or StrComp(a(0), para.item(j).translation, 1) = 0) Then
                        para.item(para.count).word = ""
                        para.item(para.count).translation = ""
                        para.item(para.count).tema = ""
                        b = True
                        End If
                Next
                If (b = False) Then
                If a(0) Like "*[а-я]*" Or a(0) Like "*ё*" Then
                    para.item(para.count).word = a(0)
                    para.item(para.count).translation = a(1)
                    para.item(para.count).tema = ws.Cells(1, i)
                    Else
                    para.item(para.count).word = a(1)
                    para.item(para.count).translation = a(0)
                    para.item(para.count).tema = ws.Cells(1, i)
                    End If
                    End If
                    b = False
                para.count = para.count + 1
                irow = irow + 1
                Wend
            irow = 2
            i = i + 1
        Wend
    
End Sub

Sub four_write(ByRef x As Test)
    Dim ws As Worksheet, ws2 As Worksheet, j As Integer
    Set ws = ThisWorkbook.Sheets("Четверки")
    Set ws2 = ThisWorkbook.Sheets("Настройки")
    Dim n As Integer, i As Integer, d As Integer
    n = x.num
    j = 0
    d = ws2.Cells(1, 1).Value
    If (d < n) Then
        n = d
        End If
    While (i < n And j < n)
    If (StrComp(x.quest_name(i).question, "", 1) = 1) Then
        ws.Cells(j + 1, 1).Value = CStr(x.quest_name(i).question) & "-" & CStr(x.quest_name(i).right)
        ws.Cells(j + 1, 2).Value = x.quest_name(i).wrong(0)
        ws.Cells(j + 1, 3).Value = x.quest_name(i).wrong(1)
        ws.Cells(j + 1, 4).Value = x.quest_name(i).wrong(2)
        j = j + 1
        End If
        i = i + 1
        Wend
End Sub

Sub TestQuest(i As Integer)
        Dim x As pairs
        Dim y As Test
        Call xls_read(x)
        Call all_four(x, y, 1)
        Call four_write(y)
        Dim j As Integer, real As Integer
    Dim ws As Worksheet, ws2 As Worksheet
        Set ws = ThisWorkbook.Sheets("Настройки")
        Set ws2 = ThisWorkbook.Sheets("Четверки")
        real = 0
        j = ws.Cells(1, 1)
        While (ws2.Cells(real + 1, 1) <> "")
            real = real + 1
            Wend
        If (j > real) Then
            j = real
        End If
        
        ws2.Cells(5, 12).Value = 0
        While (ws2.Cells(5, 12).Value < j)
        UserForm2.Show
            If (ws2.Cells(5, 12) <> j + 1) Then
            i = i + 1
            ws2.Cells(5, 12).Value = i
            End If
        Wend
        i = ws2.Cells(5, 12).Value
        If (i = j) Then
            UserForm4.Show
            End If
            ws2.Cells(5, 12).Clear
        ws2.Cells(5, 13).Clear

End Sub

Sub TestQuest_Eng(i As Integer)
        Dim x As pairs
        Dim y As Test
        Call xls_read(x)
        Call all_four(x, y, 2)
        Call four_write(y)
        Dim j As Integer, real As Integer
    Dim ws As Worksheet, ws2 As Worksheet
        Set ws = ThisWorkbook.Sheets("Настройки")
        Set ws2 = ThisWorkbook.Sheets("Четверки")
        real = 0
        j = ws.Cells(1, 1)
        While (ws2.Cells(real + 1, 1) <> "")
            real = real + 1
            Wend
        If (j > real) Then
            j = real
        End If
        
        ws2.Cells(5, 12).Value = 0
        While (ws2.Cells(5, 12).Value < j)
        UserForm3.Show
            If (ws2.Cells(5, 12) <> j + 1) Then
            i = i + 1
            ws2.Cells(5, 12).Value = i
            End If
        Wend
        i = ws2.Cells(5, 12).Value
        If (i = j) Then
            UserForm4.Show
            End If
            ws2.Cells(5, 12).Clear
        ws2.Cells(5, 13).Clear
        

End Sub

Sub MakeTest()
    UserForm1.Show
    End
End Sub

