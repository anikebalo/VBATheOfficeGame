VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Failure 
   Caption         =   "No Sale"
   ClientHeight    =   7530
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7340
   OleObjectBlob   =   "Failure.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Failure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NextDayFailure_Click() 'when moving to the next day

Dim ws As Worksheet

Set ws = Worksheets("Data")

'close the user form and update the number of clients that have been seen
Unload Me
clientnumbers = clientnumbers + 1

'reset the discount value in the data sheet to 0
ws.Range("min_40dis").Value = 0
ws.Range("min_hqdis").Value = 0
ws.Range("min_standarddis").Value = 0
ws.Range("min_carddis").Value = 0
ws.Range("min_postdis").Value = 0
ws.Range("min_envdis").Value = 0
ws.Range("min_filedis").Value = 0

'if seen less than or equal to 5 clients then move to next client
If clientnumbers <= 5 Then

    'show picture of client
    Call Clients
    
    'open first meeting user form
    DunderMifflinity.Show
    
Else 'if seen all 5 client then
    'remove client image
    Call ClearClients
    
    'show final performance analysis
    WeekEnd.Show
End If

End Sub

Private Sub percentpar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint1.Visible = True
End Sub

Private Sub profit_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint3.Visible = True
End Sub

Private Sub profitall_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint5.Visible = True
End Sub

Private Sub profitmissed_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint4.Visible = True
End Sub

Private Sub inv_lossval_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint2.Visible = True
End Sub

Private Sub missedall_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'explain value if hovered over
hint6.Visible = True
End Sub

Private Sub UserForm_Activate()

Dim ws As Worksheet

Set ws = Worksheets("Data")

'update the caption for the profit and profit missed in the userform from the data sheet
profit.Caption = Format(ws.Range("finalprice").Value, "Currency")
profitmissed.Caption = Format(ws.Range("clientmaxprice").Value, "Currency")

'show that the user missed 0% of potential profits and change color to red
percentpar.Caption = Format(0, "Percent")
percentpar.ForeColor = RGB(133, 5, 5)

'show all the profit missed and profit abtained through out the game
profitall.Caption = Format(allprofit, "Currency")
missedall.Caption = Format(allmissed, "Currency")

'update the value of unsold inventory in the userform
inv_lossval.Caption = Format(ws.Range("inv_loss").Value, "Currency")

'hide all the hints
hint1.Visible = False
hint2.Visible = False
hint3.Visible = False
hint4.Visible = False
hint5.Visible = False
hint6.Visible = False


End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

'hide all the hints when the user moves off the value
hint1.Visible = False
hint2.Visible = False
hint3.Visible = False
hint4.Visible = False
hint5.Visible = False
hint6.Visible = False
End Sub

