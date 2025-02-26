VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DunderMifflinity 
   Caption         =   "DunderMifflinity"
   ClientHeight    =   7740
   ClientLeft      =   80
   ClientTop       =   300
   ClientWidth     =   14460
   OleObjectBlob   =   "DunderMifflinity.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DunderMifflinity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Enter_Click()

Dim ctl As Control
Dim ctl2 As Control
Dim labelname As String
Dim labelnameinv As String
Dim labelctl As Control
Dim labelinvctl As Control
Dim minerror As Boolean
Dim mininverror As Boolean
Dim ws As Worksheet
Dim minimum As Long

Set ws = Worksheets("Data")

'set error flags
mininverror = True
minerror = True
        
'look through the entire user form
For Each ctl In Me.Controls
    
    'find the quantity values for each product
    If Right(ctl.Name, 1) = "q" Then
        
        'find the control name for the minimum quantity labels and the inventory lablels
        labelname = Left(ctl.Name, Len(ctl.Name) - 1)
        labelnameinv = labelname & "current"
        
        'accomodate for a difference in one of the names
        If labelname = "min_hq" Then
            labelname = "min_hqb"
        End If
        
        'finalize the control names as controls
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
    End If
Next
'''''
'inform user if flags raised
If minerror = False Then
    MsgBox "Make sure values entered are meet the client's minimum quantity requirements and are integers", vbCritical, "DunderMifflinity"
End If

'inform user if flag raised
If mininverror = False Then
    MsgBox "You do not have enough supply to fulfill this order! The maximum quantity that can be entered is the remaining inventory value", vbCritical, "DunderMifflinity"
End If

'if there are no errors
If minerror = True And mininverror = True Then

    'set quantity values in sheet to quantity value in userform
    ws.Range("min_40q").Value = min_40q.Value
    ws.Range("min_hqq").Value = min_hqq.Value
    ws.Range("min_standardq").Value = min_standardq.Value
    ws.Range("min_cardq").Value = min_cardq.Value
    ws.Range("min_postq").Value = min_postq.Value
    ws.Range("min_envq").Value = min_envq.Value
    ws.Range("min_fileq").Value = min_fileq.Value
    
    'if the quantities entered give a price below the clients budget then
    If ws.Range("missedprof") >= 0 Then
        
        'update the all profit and all missed values
        allprofit = allprofit + ws.Range("finalprice").Value
        allmissed = allmissed + ws.Range("missedprof").Value
        
        'update the inventory balance for each product
        ws.Range("min_40inv").Value = ws.Range("min_40inv").Value - CInt(min_40q.Value)
        ws.Range("min_hqinv").Value = ws.Range("min_hqinv").Value - CInt(min_hqq.Value)
        ws.Range("min_standardinv").Value = ws.Range("min_standardinv").Value - CInt(min_standardq.Value)
        ws.Range("min_cardinv").Value = ws.Range("min_cardinv").Value - CInt(min_cardq.Value)
        ws.Range("min_postinv").Value = ws.Range("min_postinv").Value - CInt(min_postq.Value)
        ws.Range("min_envinv").Value = ws.Range("min_envinv").Value - CInt(min_envq.Value)
        ws.Range("min_fileinv").Value = ws.Range("min_fileinv").Value - CInt(min_fileq.Value)
        
        'hide the form and open the congrats form
        Me.Hide
        Congrats.Show
        
        'create next clients order
        Call ClientRandom
        'clear the userform entries
        Call UserForm_Activate
        
    Else
        'inform player that their order was too expensive for the client to accept
        MsgBox "Oh no! The client did not accept that offer! It's over their budget. Time to Negotiate!", vbCritical, "DunderMifflinity"
        
        'hide the form and open the negotiation form
        Me.Hide
        Negotiation.Show
    End If

End If


End Sub

Private Sub UserForm_Activate()

'set thhe page that the form to open to
MultiPage1.Value = 1
MultiPage2.Value = 0

Dim ctl2 As Control
Dim ws As Worksheet

Set ws = Worksheets("Data")

'create the next clients order
Call ClientRandom

'update the client that is being attended
ClientNumber.Caption = clientnumbers

'update the total for this clients order to 0
running.Caption = Format(0, "Currency")
       
'make each quantity box 0
For Each ctl2 In Me.Controls
    If TypeName(ctl2) = "TextBox" Then
        ctl2.Value = 0
    End If
Next

'make the minimum quantity values the same one as in the data sheet for each product
min_40.Caption = ws.Range("min_40").Value
min_hqb.Caption = ws.Range("min_hq").Value
min_standard.Caption = ws.Range("min_standard").Value
min_card.Caption = ws.Range("min_card").Value
min_post.Caption = ws.Range("min_post").Value
min_env.Caption = ws.Range("min_env").Value
min_file.Caption = ws.Range("min_file").Value

'make the available inventory value the same one as in the data sheet for each product
min_40current.Caption = ws.Range("E3").Value
min_hqcurrent.Caption = ws.Range("E4").Value
min_standardcurrent.Caption = ws.Range("E5").Value
min_cardcurrent.Caption = ws.Range("E6").Value
min_postcurrent.Caption = ws.Range("E7").Value
min_envcurrent.Caption = ws.Range("E8").Value
min_filecurrent.Caption = ws.Range("E9").Value


End Sub
Private Sub min_40q_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity if left blank, change it to 0
If min_40q.Value = "" Then
    min_40q.Value = CInt(0)
End If

End Sub

Private Sub min_cardq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_cardq.Value = "" Then
    min_cardq.Value = CInt(0)
End If

End Sub

Private Sub min_envq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_envq.Value = "" Then
    min_envq.Value = CInt(0)
End If

End Sub

Private Sub min_fileq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_fileq.Value = "" Then
    min_fileq.Value = CInt(0)
End If

End Sub

Private Sub min_hqq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_hqq.Value = "" Then
    min_hqq.Value = CInt(0)
End If

End Sub

Private Sub min_postq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_postq.Value = "" Then
    min_postq.Value = CInt(0)
End If

End Sub

Private Sub min_standardq_AfterUpdate()

'update the total for this clients order
Call AfterChangeOG

'if a quantity is left blank, change it to 0
If min_standardq.Value = "" Then
    min_standardq.Value = CInt(0)
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

Private Sub UserForm_Initialize()

Dim ws As Worksheet

Set ws = Worksheets("Data")

'set the profit, profitmissed and client number values
allprofit = 0
allmissed = 0
clientnumbers = 1

'get the client pictures
Call Clients

'set the quantity values in the userform
min_40q.Value = 0
min_hqq.Value = 0
min_standardq.Value = 0
min_cardq.Value = 0
min_postq.Value = 0
min_envq.Value = 0
min_fileq.Value = 0

'set the discounts in the data sheet
ws.Range("min_40dis").Value = 0
ws.Range("min_hqdis").Value = 0
ws.Range("min_standarddis").Value = 0
ws.Range("min_carddis").Value = 0
ws.Range("min_postdis").Value = 0
ws.Range("min_envdis").Value = 0
ws.Range("min_filedis").Value = 0

'set this orders total value to 0
running.Caption = Format(0, "Currency")

'randomly assign inventory quantities for each product
Call InvRandom

'update inventory value labels in the userform to the values in the data sheet
inv_40 = Worksheets("Data").Range("E3").Value
inv_hq = Worksheets("Data").Range("E4").Value
inv_standard = Worksheets("Data").Range("E5").Value
inv_card = Worksheets("Data").Range("E6").Value
inv_post = Worksheets("Data").Range("E7").Value
inv_env = Worksheets("Data").Range("E8").Value
inv_file = Worksheets("Data").Range("E9").Value

'remind user that they must read all the instructions
MsgBox "Ensure that you have read *ALL* the instructions on the first tab before going to the First Meeting tab!", vbInformation, "DunderMifflinity"

'start userform with page 1 for both multipages
MultiPage1.Value = 0
MultiPage2.Value = 0

End Sub
