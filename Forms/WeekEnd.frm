VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WeekEnd 
   Caption         =   "Performance Evaluation"
   ClientHeight    =   11800
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7270
   OleObjectBlob   =   "WeekEnd.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "WeekEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Thanks_Click()

'Close the userform
Unload Me
Unload DunderMifflinity

End Sub

Private Sub UserForm_Activate()

Dim ws As Worksheet
Dim PARcalc As Double
Dim invloss As Double

Set ws = Worksheets("Data")

'remaining value of unsold inventory
invloss = ws.Range("inv_loss").Value

'display final profit
profitall.Caption = Format(allprofit, "Currency")

'display final missed profit
missedall.Caption = Format(allmissed, "Currency")

'update userform with value of unsold inventory
invlossval.Caption = Format(invloss, "Currency")

'make sure that relevant values are greater than 0
If allprofit > 0 And allmissed > 0 Then
    'calculate percent of potential profits achieved, including value of unsold inventory
    PARcalc = allprofit / (allprofit + allmissed + invloss)
    'place value in userform
    PAR.Caption = Format(PARcalc, "Percent")
Else
    PAR.Caption = "-"
End If

'if the user did well (PAR >= 0.7) change the text and update colour to green
If PARcalc >= 0.7 Then
    employstatus.Caption = "You keep your job!"
    employstatus.ForeColor = RGB(8, 132, 4)
Else
'if the user did not do well (PAR < 0.7) change the text and update colour to red
    employstatus.Caption = "Sorry, but you're fired"
    employstatus.ForeColor = RGB(200, 4, 4)
End If
    

End Sub

