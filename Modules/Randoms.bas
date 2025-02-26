Attribute VB_Name = "Randoms"
Option Explicit

Public Sub ClientRandom()

Dim min_counterrandom As Long
Dim min_randompick As Integer
Dim max_counterrandom As Long
Dim max_randompick As Integer
Dim i As Integer
Dim ws As Worksheet

Set ws = Worksheets("Data")

'set row counter
i = 3

'set counters
min_counterrandom = 0
max_counterrandom = 0

'create random min and max client expectations for each product
Do Until min_counterrandom = 7 '7 products available
    min_randompick = Int(9 * Rnd) 'the minimum number of the product the client wants
    max_randompick = Application.WorksheetFunction.RandBetween(min_randompick + 1, min_randompick + 5) 'the maximum must be 1-5 greater than the minimum
    ws.Cells(i, 4).Value = max_randompick 'place the max (only used for max price)
    ws.Cells(i, 3).Value = min_randompick 'place the min
    
    'increase the counters
    i = i + 1
    min_counterrandom = min_counterrandom + 1
    max_counterrandom = max_counterrandom + 1
Loop


End Sub

Public Sub HintRandom()

Dim ws As Worksheet

Set ws = Worksheets("Data")

'create a hint value between 70% and 100% of the clients max price
ws.Range("hint_random").Value = Rnd() * 0.3 + 0.7


End Sub

Public Sub InvRandom() ' assign inventory to each product

Dim counterrandom As Long
Dim randomquant As Integer
Dim i As Integer
Dim ws As Worksheet

Set ws = Worksheets("Data")

'set row counter
i = 3

'set counters
counterrandom = 0

'create random min and max client expectations for each product
Do Until counterrandom = 7 '7 products available
    randomquant = Application.WorksheetFunction.RandBetween(25, 40) 'the inventory quantity for each product must be between 25 and 40
    ws.Cells(i, 5).Value = randomquant 'inventory quantity in the right row
    
    'increase the counters
    i = i + 1
    counterrandom = counterrandom + 1
Loop

End Sub

Public Sub DiscountRandom() 'assign max discount to each product

Dim DiscountRandom As Long
Dim discountcounter As Integer
Dim i As Integer
Dim cell As Range
Dim ws As Worksheet

Set ws = Worksheets("Data")

'set counters
i = 15 'row counter
discountcounter = 0

'look at the entered discount, if its not empty, assign a maximum discount value (between 5% and 70%)
For Each cell In ws.Range("B15:B21")
    If cell.Value <> 0 Or cell.Value <> "" Then
        cell.Offset(0, 3).Value = Rnd * (0.05 - 0.7) + 0.7
    Else: cell.Offset(0, 3).ClearContents
    End If
Next cell

End Sub


Sub StartDunderFinity()
'open the dundermifflinity userform
DunderMifflinity.Show

End Sub
