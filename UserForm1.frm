VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "English Tutor"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bal As Double
Dim nVopros As Double
Private Sub Label1_Click()

End Sub

Private Sub OptionButton1_Click()
    Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Четверки")
    Unload Me
    Dim i As Integer
    i = ws.Cells(5, 12)
    Call TestQuest(i)
End Sub

Private Sub OptionButton2_Click()
    Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Четверки")
    Unload Me
    Dim i As Integer
    i = ws.Cells(5, 12)
    Call TestQuest_Eng(i)
End Sub
Private Sub UserForm_Activate()

End Sub

