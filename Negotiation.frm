VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Negotiation 
   Caption         =   "Negotiation"
   ClientHeight    =   8380.001
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   14360
   OleObjectBlob   =   "Negotiation.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Negotiation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Enter_Click()

Dim ctl As Control
Dim ctl2 As Control
Dim ctl3 As Control
Dim labelnameinv As String
Dim labelname As String
Dim labelctl As Control
Dim labelname2 As String
Dim labelinvctl As Control
Dim labelcontrol2 As Control
Dim minerror As Boolean
Dim minpercenterror As Boolean
Dim mininverror As Boolean
Dim minimum As Long
Dim ws As Worksheet

Set ws = Worksheets("Data")

'set error flags
minpercenterror = True
minerror = True
mininverror = True

'look through the userform
For Each ctl In Me.Controls

    'look for the quantity value
    If Right(ctl.Name, 1) = "q" Then
        labelname = Left(ctl.Name, Len(ctl.Name) - 1)
        labelnameinv = labelname & "current"
        
        'accomodate for a difference in one of the names
        If labelname = "min_hq" Then
            labelname = "min_hqb"
        End If
        
        'find the control name for the minimum quantity labels and the inventory lablels
        Set labelctl = Me.Controls(labelname)
        Set labelinvctl = Me.Controls(labelnameinv)
        
        'find the minimum amount of inventory available (either the minimum quantity needed to purchase or the inventory remaning)
        minimum = Application.WorksheetFunction.Min(CDbl(labelctl.Caption), CDbl(labelinvctl.Caption))
        
        'ensure that quantity entered is numeric
        If Not IsNumeric(ctl.Value) Then
            ctl.BackColor = RGB(255, 0, 0)
            minerror = False
            
        'ensure that user can only purchase as much inventory as available and they are entering positive integers
        ElseIf CDbl(ctl.Value) < minimum Or CDbl(ctl.Value) < 0 Or CInt(ctl.Value) <> ctl.Value Then
            ctl.BackColor = RGB(255, 0, 0)
            minerror = False
            
        'ensure that they are not ordering more than is available
        ElseIf CDbl(labelinvctl.Caption) - ctl.Value < 0 Then
            ctl.BackColor = RGB(255, 0, 0)
            mininverror = False
        End If
    
    'find the discount value
    ElseIf Right(ctl.Name, 3) = "dis" Then
        
        'if its not a number then flag as error and change colour to red
        If Not IsNumeric(ctl.Value) Then
            minpercenterror = False
            ctl.BackColor = RGB(255, 0, 0)
        
        'if its negative or greater than 0.7 then flag as error and change colour to red
        ElseIf CDbl(ctl.Value) < 0 Or CDbl(ctl.Value) > 0.7 Then
            minpercenterror = False
            ctl.BackColor = RGB(255, 0, 0)
        End If
    End If
Next

'inform user of errors
If minerror = False And minpercenterror = False Then
    MsgBox "Make sure values entered meet the client's minimum quanitity requirements and are integers" & Chr(10) & _
    "AND Make sure that all discounts are less than or equal to 0.7 (70%)", vbCritical, "DunderMifflinity"
ElseIf minerror = False Then
    MsgBox "Make sure values entered meet the client's minimum quantity requirements", vbCritical, "DunderMifflinity"
ElseIf minpercenterror = False Then
    MsgBox "Make sure that all discounts are less than or equal to 0.7 (70%)", vbCritical, "DunderMifflinity"
ElseIf mininverror = False Then
    MsgBox "You do not have enough supply to fulfill this order! The maximum quantity that can be entered is the remaining inventory value", vbCritical, "DunderMifflinity"
    
'if there are no errors then
Else:
    'put the inputed quantities in the data sheet
    ws.Range("min_40q").Value = min_40q.Value
    ws.Range("min_hqq").Value = min_hqq.Value
    ws.Range("min_standardq").Value = min_standardq.Value
    ws.Range("min_cardq").Value = min_cardq.Value
    ws.Range("min_postq").Value = min_postq.Value
    ws.Range("min_envq").Value = min_envq.Value
    ws.Range("min_fileq").Value = min_fileq.Value
    
    'put the inputted discounts in the data sheet
    ws.Range("min_40dis").Value = min_40dis.Value
    ws.Range("min_hqdis").Value = min_hqdis.Value
    ws.Range("min_standarddis").Value = min_standarddis.Value
    ws.Range("min_carddis").Value = min_carddis.Value
    ws.Range("min_postdis").Value = min_postdis.Value
    ws.Range("min_envdis").Value = min_envdis.Value
    ws.Range("min_filedis").Value = min_filedis.Value
        
    'look through the data shee
    For Each ctl3 In Me.Controls
        'for each discount value
        
        If Right(ctl3.Name, 3) = "dis" Then
        
            'if a discount was entered
            If ctl3.Value <> 0 Then
            
                'find the labelname for the approval label in the userform (indicating whether or not your manager accepted the discount)
                labelname2 = Left(ctl3.Name, Len(ctl3.Name) - 3) & "approv"
                
                'set the the labelname as a control
                Set labelcontrol2 = Me.Controls(labelname2)
                
                'look in the sheet and change the approval label. if the sheet says it can be accepted (1), change to green an say yes
                If ws.Range(ctl3.Name).Offset(, 1) = 1 Then
                    labelcontrol2.BackColor = RGB(0, 255, 0)
                    labelcontrol2.Caption = "Yes"
                Else
                'look in the sheet and change the approval label. if the sheet says it CANNOT be accepted (0), change to red an say no
                    labelcontrol2.BackColor = RGB(255, 0, 0)
                    labelcontrol2.Caption = "No"
                End If
                'show the approval label
                labelcontrol2.Visible = True
            End If
        End If
    Next
    
    'stop for 2 seconds to give users a chance to see manager feedback
    Application.Wait (Now + TimeValue("0:00:02"))
    
    'if the quantities entered give a price below the clients budget then
    If ws.Range("missedprof") >= 0 Then
    
        'update the all profit and all missed values
        allprofit = allprofit + ws.Range("finalprice").Value
        allmissed = allmissed + ws.Range("missedprof").Value
        
        ''update the inventory balance for each product
        ws.Range("min_40inv").Value = ws.Range("min_40inv").Value - CInt(min_40q.Value)
        ws.Range("min_hqinv").Value = ws.Range("min_hqinv").Value - CInt(min_hqq.Value)
        ws.Range("min_standardinv").Value = ws.Range("min_standardinv").Value - CInt(min_standardq.Value)
        ws.Range("min_cardinv").Value = ws.Range("min_cardinv").Value - CInt(min_cardq.Value)
        ws.Range("min_postinv").Value = ws.Range("min_postinv").Value - CInt(min_postq.Value)
        ws.Range("min_envinv").Value = ws.Range("min_envinv").Value - CInt(min_envq.Value)
        ws.Range("min_fileinv").Value = ws.Range("min_fileinv").Value - CInt(min_fileq.Value)
        
        'reset this orders total value
        running.Caption = Format(0, "Currency")
        
        'reset the all quantities and discounts
        For Each ctl2 In Me.Controls
            If TypeName(ctl2) = "TextBox" Then
                ctl2.Value = 0
            End If
        Next
        
        
        'hide the form and open the congrats form
        Unload Me
        Congrats.Show
        
        'create new client order
        Call ClientRandom
    Else
        'if the negotiation fails the first time then
        If negotiationcount = 1 Then
        
            'display the last price the client was charged
            lastprice.Caption = Format(ws.Range("finalprice").Value, "Currency")
            
            'remember that this is clients 2nd try & update the label in the userform
            negotiationcount = 2
            negotiation_num.Caption = negotiationcount
            'inform the user
            MsgBox "The client rejected your offer, you have one more chance to change their minds!", vbCritical, "DunderMifflinity"
            
            'update the total order price with the last charged price
            running.Caption = Format(ws.Range("finalprice").Value, "Currency")
            
            'give the user the option to have a hint
            HintButton.Visible = True
            
            'change the discounts back to 0 in the data sheet
            ws.Range("min_40dis").Value = 0
            ws.Range("min_hqdis").Value = 0
            ws.Range("min_standarddis").Value = 0
            ws.Range("min_carddis").Value = 0
            ws.Range("min_postdis").Value = 0
            ws.Range("min_envdis").Value = 0
            ws.Range("min_filedis").Value = 0
        
        Else:
            'change the product quantities to 0 in the data sheet
            ws.Range("min_40q").Value = 0
            ws.Range("min_hqq").Value = 0
            ws.Range("min_standardq").Value = 0
            ws.Range("min_cardq").Value = 0
            ws.Range("min_postq").Value = 0
            ws.Range("min_envq").Value = 0
            ws.Range("min_fileq").Value = 0
            
            'update the allmissed value
            allmissed = allmissed + ws.Range("clientmaxprice").Value
            
            'close the user form & open failure form
            Unload Me
            Failure.Show
            
            'create new random orders for next client
            Call ClientRandom
        End If
    End If
    
End If
End Sub

Private Sub HintButton_Click() 'if the user asks for a hint

'generate a new hint value
Call HintRandom

'open the hints user form
Hints.Show

'hind the hints button
HintButton.Visible = False

End Sub

Public Sub UserForm_Activate()

'when the form opens, set multipage 1 and 2 to open default pages
MultiPage1.Value = 1
MultiPage2.Value = 0

Dim ws As Worksheet

Set ws = Worksheets("Data")

'update the minimum quantity values into the userform from the sheet
min_40.Caption = ws.Range("min_40").Value
min_hqb.Caption = ws.Range("min_hq").Value
min_standard.Caption = ws.Range("min_standard").Value
min_card.Caption = ws.Range("min_card").Value
min_post.Caption = ws.Range("min_post").Value
min_env.Caption = ws.Range("min_env").Value
min_file.Caption = ws.Range("min_file").Value

'update the inventory values into the userform from the sheet
min_40current.Caption = ws.Range("E3").Value
min_hqcurrent.Caption = ws.Range("E4").Value
min_standardcurrent.Caption = ws.Range("E5").Value
min_cardcurrent.Caption = ws.Range("E6").Value
min_postcurrent.Caption = ws.Range("E7").Value
min_envcurrent.Caption = ws.Range("E8").Value
min_filecurrent.Caption = ws.Range("E9").Value

'change discount values to 0 in the sheet
ws.Range("min_40dis").Value = 0
ws.Range("min_hqdis").Value = 0
ws.Range("min_standarddis").Value = 0
ws.Range("min_carddis").Value = 0
ws.Range("min_postdis").Value = 0
ws.Range("min_envdis").Value = 0
ws.Range("min_filedis").Value = 0

'hide all approval labels
min_40approv.Visible = False
min_hqapprov.Visible = False
min_standardapprov.Visible = False
min_cardapprov.Visible = False
min_postapprov.Visible = False
min_envapprov.Visible = False
min_fileapprov.Visible = False

'start at negotiation 1 an update label in the userform
negotiationcount = 1
ClientNumber.Caption = clientnumbers

'hide the hint button
HintButton.Visible = False

'update the negotiation label from negotiation count
negotiation_num.Caption = negotiationcount

'update the lastprice caption from the last price charged in the DunderMifflinity form
lastprice.Caption = DunderMifflinity.running.Caption

End Sub
Private Sub UserForm_Initialize()

'set userform quantities to 0
min_40q.Value = 0
min_hqq.Value = 0
min_standardq.Value = 0
min_cardq.Value = 0
min_postq.Value = 0
min_envq.Value = 0
min_fileq.Value = 0

'set userform total price charged client to 0
running.Caption = Format(0, "Currency")

'set discount values to 0
min_40dis.Value = 0
min_hqdis.Value = 0
min_standarddis.Value = 0
min_carddis.Value = 0
min_postdis.Value = 0
min_envdis.Value = 0
min_filedis.Value = 0

'remind user to read instructions
MsgBox "Ensure that you have read *ALL* the instructions on the first tab before going to the Negotiations tab!", vbInformation, "DunderMifflinity"

End Sub
Private Sub min_40q_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_40q.Value = "" Then
    min_40q.Value = CInt(0)
End If


End Sub

Private Sub min_40dis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_40dis.Value = "" Then
    min_40dis.Value = CInt(0)
End If

End Sub

Private Sub min_cardq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_cardq.Value = "" Then
    min_cardq.Value = CInt(0)
End If


End Sub

Private Sub min_carddis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_carddis.Value = "" Then
    min_carddis.Value = CInt(0)
End If


End Sub

Private Sub min_envq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_envq.Value = "" Then
    min_envq.Value = CInt(0)
End If


End Sub

Private Sub min_envdis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_envdis.Value = "" Then
    min_envdis.Value = CInt(0)
End If


End Sub

Private Sub min_fileq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_fileq.Value = "" Then
    min_fileq.Value = CInt(0)
End If


End Sub

Private Sub min_filedis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_filedis.Value = "" Then
    min_filedis.Value = CInt(0)
End If


End Sub

Private Sub min_hqq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_hqq.Value = "" Then
    min_hqq.Value = CInt(0)
End If


End Sub

Private Sub min_hqdis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_hqdis.Value = "" Then
    min_hqdis.Value = CInt(0)
End If


End Sub

Private Sub min_postq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_postq.Value = "" Then
    min_postq.Value = CInt(0)
End If


End Sub

Private Sub min_postdis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_postdis.Value = "" Then
    min_postdis.Value = CInt(0)
End If


End Sub

Private Sub min_standardq_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_standardq.Value = "" Then
    min_standardq.Value = CInt(0)
End If

End Sub

Private Sub min_standarddis_AfterUpdate()

'update the total for this clients order
Call AfterChangeNego

'if a quantity is left blank, change it to 0
If min_standarddis.Value = "" Then
    min_standarddis.Value = CInt(0)
End If


End Sub

Private Sub min_40q_Change()

'every time a quantity is changed, make the background color white
min_40q.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_hqq_Change()
'every time a quantity is changed, make the background color white
min_hqq.BackColor = RGB(255, 255, 255)

End Sub


Private Sub min_standardq_Change()
'every time a quantity is changed, make the background color white
min_standardq.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_cardq_Change()
'every time a quantity is changed, make the background color white
min_cardq.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_postq_Change()
'every time a quantity is changed, make the background color white
min_postq.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_envq_Change()
'every time a quantity is changed, make the background color white
min_envq.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_fileq_Change()
'every time a quantity is changed, make the background color white
min_fileq.BackColor = RGB(255, 255, 255)

End Sub


Private Sub min_40dis_Change()
'every time a quantity is changed, make the background color white
min_40dis.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_hqdis_Change()
'every time a quantity is changed, make the background color white
min_hqdis.BackColor = RGB(255, 255, 255)

End Sub


Private Sub min_standarddis_Change()
'every time a quantity is changed, make the background color white
min_standarddis.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_carddis_Change()
'every time a quantity is changed, make the background color white
min_carddis.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_postdis_Change()
'every time a quantity is changed, make the background color white
min_postdis.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_envdis_Change()
'every time a quantity is changed, make the background color white
min_envdis.BackColor = RGB(255, 255, 255)

End Sub

Private Sub min_filedis_Change()
'every time a quantity is changed, make the background color white
min_filedis.BackColor = RGB(255, 255, 255)

End Sub


