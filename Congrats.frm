VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Congrats 
   Caption         =   "Congrats! A Sale!"
   ClientHeight    =   7340
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   7350
   OleObjectBlob   =   "Congrats.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Congrats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub NextDayCongrats_Click() 'once the user clicks Next Day

Dim ws As Worksheet

Set ws = Worksheets("Data")

'close the userform
Unload Me

'mark that process is moving to the next client
clientnumbers = clientnumbers + 1

'reset discount values in data sheet to 0
ws.Range("min_40dis").Value = 0
ws.Range("min_hqdis").Value = 0
ws.Range("min_standarddis").Value = 0
ws.Range("min_carddis").Value = 0
ws.Range("min_postdis").Value = 0
ws.Range("min_envdis").Value = 0
ws.Range("min_filedis").Value = 0

'move to the next client if user hasn't seen all 5 clients yet
If clientnumbers <= 5 Then
    Call Clients 'show picture of client
    Call DiscountRandom 'set max discounts fot next round of play
    DunderMifflinity.Show 'open entry form
    
'if user has seen 5 clients then remove last client picture
Else
    Call ClearClients
    'show final performance
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
Dim PARcalc As Double
Dim soldprice As Double
Dim availprice As Double
Dim missedprice As Double

Set ws = Worksheets("Data")

'set values based on datasheet
soldprice = CDbl(ws.Range("finalprice").Value) 'how much the client bought the products for
missedprice = CDbl(ws.Range("missedprof").Value) 'how much more than the sold price they were willing to spend
availprice = CDbl(ws.Range("clientmaxprice").Value) 'how much they were willing to spend

profit.Caption = Format(soldprice, "Currency") 'place sold price on userform
profitmissed.Caption = Format(missedprice, "Currency") 'place profit missed on userform

PARcalc = soldprice / availprice 'calculate percent of available rev that was realized

'if its greater or equal to 0.7, make it green, otherwise red
If PARcalc >= 0.7 Then
    percentpar.ForeColor = RGB(0, 128, 0)
Else
    percentpar.ForeColor = RGB(135, 5, 5)
End If

'show the percent of available revenue
percentpar.Caption = Format(PARcalc, "Percent")

'show the profit and profit missed
profitall.Caption = Format(allprofit, "Currency")
missedall.Caption = Format(allmissed, "Currency")

'show the value of unsold inventory left
inv_lossval.Caption = Format(ws.Range("inv_loss").Value, "Currency")

'do not show the hints
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
