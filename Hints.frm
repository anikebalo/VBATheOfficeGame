VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Hints 
   Caption         =   "Hint Required"
   ClientHeight    =   4630
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7510
   OleObjectBlob   =   "Hints.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Hints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BacktoNego_Click()

Dim confirm As String

'confirm with user that they know the consequences of closing the negotiations form
confirm = MsgBox("Are you sure you're ready to leave the hint box? Once you leave, the hint will go away for good!", vbYesNo, "DunderMifflinity")

'if they confirm, then close the form, otherwise leave open
If confirm = vbYes Then
    Unload Me
End If

End Sub

Private Sub UserForm_Activate()

Dim ws As Worksheet

Set ws = Worksheets("Data")

'get the hint value from the data sheet
hintvalue.Caption = Format(ws.Range("hintval").Value, "Currency")

End Sub

