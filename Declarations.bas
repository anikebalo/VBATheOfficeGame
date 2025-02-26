Attribute VB_Name = "Declarations"
Option Explicit

Public allprofit As Double
Public allmissed As Double
Public clientnumbers As Integer
Public negotiationcount As Integer
Public inv_40 As Integer
Public inv_hq As Integer
Public inv_standard As Integer
Public inv_card As Integer
Public inv_post As Integer
Public inv_env As Integer
Public inv_file As Integer

Public Sub AfterChangeOG() 'update subtotal after updated in userform

Dim gochange As Boolean
Dim total As Double

'ensure that all values are numeric and positive
If IsNumeric(DunderMifflinity.min_40q.Value) And DunderMifflinity.min_40q.Value >= 0 And IsNumeric(DunderMifflinity.min_hqq.Value) And DunderMifflinity.min_hqq.Value >= 0 And IsNumeric(DunderMifflinity.min_standardq.Value) _
And DunderMifflinity.min_standardq.Value >= 0 And IsNumeric(DunderMifflinity.min_cardq.Value) And DunderMifflinity.min_cardq.Value >= 0 And IsNumeric(DunderMifflinity.min_postq.Value) And DunderMifflinity.min_postq.Value >= 0 And IsNumeric(DunderMifflinity.min_envq.Value) _
And DunderMifflinity.min_envq.Value >= 0 And IsNumeric(DunderMifflinity.min_fileq.Value) And DunderMifflinity.min_fileq.Value >= 0 Then
    gochange = True
Else
    gochange = False
End If

'if all are numeric then, multiply quantity entered by product price
If gochange = True Then
    total = CDbl(DunderMifflinity.min_40q.Value) * 25 + CDbl(DunderMifflinity.min_hqq.Value) * 12 + _
            CDbl(DunderMifflinity.min_standardq.Value) * 10 + CDbl(DunderMifflinity.min_cardq.Value) * 45 + _
            CDbl(DunderMifflinity.min_postq.Value) * 8 + CDbl(DunderMifflinity.min_envq.Value) * 15 + _
            CDbl(DunderMifflinity.min_fileq.Value) * 10
    'update the total order price for this client in the userform
    DunderMifflinity.running.Caption = Format(total, "Currency")
        
        'if there is an error, change the total order price to 0
Else: DunderMifflinity.running.Caption = Format(0, "Currency")
End If

End Sub

Public Sub AfterChangeNego()

Dim gochange As Boolean
Dim total As Double
Dim gopercent As Boolean

'ensure that all values are numeric and positive
If IsNumeric(Negotiation.min_40q.Value) And Negotiation.min_40q.Value >= 0 And _
    IsNumeric(Negotiation.min_hqq.Value) And Negotiation.min_hqq.Value >= 0 And _
    IsNumeric(Negotiation.min_standardq.Value) And Negotiation.min_standardq.Value >= 0 And _
    IsNumeric(Negotiation.min_cardq.Value) And Negotiation.min_cardq.Value >= 0 And _
    IsNumeric(Negotiation.min_postq.Value) And Negotiation.min_postq.Value >= 0 And _
    IsNumeric(Negotiation.min_envq.Value) And Negotiation.min_envq.Value >= 0 And _
    IsNumeric(Negotiation.min_fileq.Value) And Negotiation.min_fileq.Value >= 0 Then
    gochange = True
Else
    gochange = False
End If

'ensure that all values are numeric and positive and less than or equal to 0.7
If IsNumeric(Negotiation.min_40dis.Value) And Negotiation.min_40dis.Value >= 0 And Negotiation.min_40dis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_hqdis.Value) And Negotiation.min_hqdis.Value >= 0 And Negotiation.min_hqdis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_standarddis.Value) And Negotiation.min_standarddis.Value >= 0 And Negotiation.min_standarddis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_carddis.Value) And Negotiation.min_carddis.Value >= 0 And Negotiation.min_carddis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_postdis.Value) And Negotiation.min_postdis.Value >= 0 And Negotiation.min_postdis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_envdis.Value) And Negotiation.min_envdis.Value >= 0 And Negotiation.min_envdis.Value <= 0.7 _
    And IsNumeric(Negotiation.min_filedis.Value) And Negotiation.min_filedis.Value >= 0 And Negotiation.min_filedis.Value <= 0.7 Then
        gopercent = True
Else
        gopercent = False
End If

'if no errors then update total order price for the client in the userform
If gochange = True And gopercent = True Then
    total = (CDbl(Negotiation.min_40q.Value) * 25 * CDbl(1 - Negotiation.min_40dis.Value)) + (CDbl(Negotiation.min_hqq.Value) * 12 * CDbl(1 - Negotiation.min_hqdis.Value)) + _
            (CDbl(Negotiation.min_standardq.Value) * 10 * CDbl(1 - Negotiation.min_standarddis.Value)) + (CDbl(Negotiation.min_cardq.Value) * 45 * CDbl(1 - Negotiation.min_carddis.Value)) + _
            (CDbl(Negotiation.min_postq.Value) * 8 * CDbl(1 - Negotiation.min_postdis.Value)) + (CDbl(Negotiation.min_envq.Value) * 15 * CDbl(1 - Negotiation.min_envdis.Value)) + _
            (CDbl(Negotiation.min_fileq.Value) * 10 * CDbl(1 - Negotiation.min_filedis.Value))
    'update the total order price for this client in the userform
    Negotiation.running.Caption = Format(total, "Currency")
    
        'if there is an error, change the total order price to 0
Else: Negotiation.running.Caption = Format(0, "Currency")
End If


End Sub
