VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "English Tutor"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bal As Integer
Dim n As Integer
Dim x As pairs
Dim bool(3) As Boolean
Dim y As Test
Dim ans As Integer

Private Sub CommandButton1_Click()

    Dim ws As Worksheet, ws1 As Worksheet
    Set ws = ThisWorkbook.Sheets("Четверки")
    Dim n As Integer
    Dim count As Integer
    count = ws.Cells(5, 12).Value
    Dim a() As String
    a = Split(ws.Cells(count + 1, 1), "-")
    i = 0
    If (OptionButton1.Value = True And ans = 1) Then
    n = ws.Cells(5, 13).Value
    ws.Cells(5, 13).Value = (n + 1)
        Unload Me
    
    ElseIf (OptionButton2.Value = True And ans = 2) Then
    n = ws.Cells(5, 13).Value
    ws.Cells(5, 13).Value = (n + 1)
        Unload Me
    
    ElseIf (OptionButton3.Value = True And ans = 3) Then
    n = ws.Cells(5, 13).Value
    ws.Cells(5, 13).Value = (n + 1)
        Unload Me
    
    ElseIf (OptionButton4.Value = True And ans = 4) Then
    n = ws.Cells(5, 13).Value
    ws.Cells(5, 13).Value = (n + 1)
        Unload Me
    End If

    If (OptionButton1.Value = True And ans <> 1) Then
        If (ans = 2) Then
            MsgBox (CStr(OptionButton2.Caption))
        ElseIf (ans = 3) Then
            MsgBox (CStr(OptionButton3.Caption))
        Else: MsgBox (CStr(OptionButton4.Caption))
        End If
    ElseIf (OptionButton2.Value = True And ans <> 2) Then
        If (ans = 1) Then
            MsgBox (CStr(OptionButton1.Caption))
        ElseIf (ans = 3) Then
            MsgBox (CStr(OptionButton3.Caption))
        Else: MsgBox (CStr(OptionButton4.Caption))
        End If
    ElseIf (OptionButton3.Value = True And ans <> 3) Then
        If (ans = 1) Then
            MsgBox (CStr(OptionButton1.Caption))
        ElseIf (ans = 2) Then
            MsgBox (CStr(OptionButton2.Caption))
        Else: MsgBox (CStr(OptionButton4.Caption))
        End If
    ElseIf (OptionButton4.Value = True And ans <> 4) Then
        If (ans = 1) Then
            MsgBox (CStr(OptionButton1.Caption))
        ElseIf (ans = 2) Then
            MsgBox (CStr(OptionButton2.Caption))
        Else: MsgBox (CStr(OptionButton3.Caption))
        End If
    ElseIf (OptionButton1.Value = False And OptionButton2.Value = False And OptionButton3.Value = False And OptionButton4.Value = False) Then
        If (ans = 1) Then
            MsgBox (CStr(OptionButton1.Caption))
        ElseIf (ans = 2) Then
            MsgBox (CStr(OptionButton2.Caption))
        ElseIf (ans = 3) Then
            MsgBox (CStr(OptionButton3.Caption))
        Else: MsgBox (CStr(OptionButton4.Caption))
        End If
    End If

   
    Unload Me
End Sub

Private Sub CommandButton2_Click()
    Dim ws As Worksheet, ws1 As Worksheet, ws3 As Worksheet
    Set ws = ThisWorkbook.Sheets("Четверки")
    Set ws1 = ThisWorkbook.Sheets("Настройки")
    Set ws3 = ThisWorkbook.Sheets("Слова и группы")
    Dim x As pairs, y As Test, d As Integer
    Call xls_read(x)
    Call all_four(x, y, 2)
    Dim n As Integer
    n = ws1.Cells(1, 1).Value
    While (ws.Cells(d + 1, 1) <> "")
    d = d + 1
    Wend
    n = ws1.Cells(1, 1).Value
    If (n > d) Then
    n = d
    End If
    ws.Cells(5, 12).Value = (n + 1)
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub OptionButton4_Click()

End Sub

Private Sub UserForm_Activate()
    Dim ws As Worksheet, ws1 As Worksheet, ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("Слова и группы")
    Set ws1 = ThisWorkbook.Sheets("Четверки")
        Set ws = ThisWorkbook.Sheets("Настройки")
        n = ws.Cells(1, 1)
    Call xls_read(x)
    Dim count As Integer
    count = ws1.Cells(5, 12).Value
    'Call four_gen1(x, y, 1, 2, 1)

    ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(1) & "," & CStr(4) & ")"
    n = ws.Cells(5, 10).Value
    ws.Cells(5, 10).Clear
    Dim i As Integer
    Dim a() As String
    a = Split(ws1.Cells(count + 1, 1), "-")
    Label1.Caption = a(1)
    i = 0
    If (n = 1) Then
        OptionButton1.Caption = a(0)
        ans = 1
        bool(n - 1) = True
        While (i < 3)
        While (bool(n - 1) = True)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(1) & "," & CStr(4) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            If (n = 2) Then
                OptionButton2.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 3) Then
                OptionButton3.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 4) Then
                OptionButton4.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            End If
            i = i + 1
            Wend
    ElseIf (n = 2) Then
        OptionButton2.Caption = a(0)
        ans = 2
        bool(n - 1) = True
        While (i < 3)
        While (bool(n - 1) = True)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(1) & "," & CStr(4) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            If (n = 1) Then
                OptionButton1.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 3) Then
                OptionButton3.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 4) Then
                OptionButton4.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            End If
            i = i + 1
            Wend
    ElseIf (n = 3) Then
        OptionButton3.Caption = a(0)
        ans = 3
        bool(n - 1) = True
        While (i < 3)
        While (bool(n - 1) = True)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(1) & "," & CStr(4) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            If (n = 1) Then
                OptionButton1.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 2) Then
                OptionButton2.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 4) Then
                OptionButton4.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            End If
            i = i + 1
            Wend
    Else
        OptionButton4.Caption = a(0)
        ans = 4
        bool(n - 1) = True
        While (i < 3)
        While (bool(n - 1) = True)
            ws.Cells(5, 10).Value = "=RANDBETWEEN(" & CStr(1) & "," & CStr(4) & ")"
            n = ws.Cells(5, 10).Value
            ws.Cells(5, 10).Clear
            Wend
            If (n = 1) Then
                OptionButton1.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 2) Then
                OptionButton2.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            ElseIf (n = 3) Then
                OptionButton3.Caption = ws1.Cells(count + 1, i + 2)
                bool(n - 1) = True
            End If
            i = i + 1
            Wend
    End If
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        Call CommandButton2_Click
        Me.Hide
    End If
End Sub

