VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "English Tutor"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()
End Sub

Private Sub UserForm_Activate()
    Label2.Caption = "Тестирование закончено."
    Dim ws As Worksheet, ws1 As Worksheet
    Set ws = ThisWorkbook.Sheets("Четверки")
    Set ws1 = ThisWorkbook.Sheets("Настройки")
    Dim n As Integer, m As Integer, real As Integer
    real = 0
        While (ws.Cells(real + 1, 1) <> "")
            real = real + 1
            Wend
            
    n = ws.Cells(5, 13).Value
    m = ws1.Cells(1, 1).Value
    If (m > real) Then
        m = real
        End If
    Label1.Caption = CStr(n) & " правильных ответов" & " из " & CStr(m)
End Sub


