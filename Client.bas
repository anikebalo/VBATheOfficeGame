Attribute VB_Name = "Client"
Option Explicit
Public Sub Clients()

Dim sh_current As Shape
Dim sh As Shape
Dim sh_paste As Shape

Dim ws As Worksheet
Dim wsclient As Worksheet

Set ws = Worksheets("DunderMifflin")
Set wsclient = Worksheets("clients")

For Each sh In ws.Shapes
    If sh.Name = "CurrentClient" Then
        sh.Delete
        Exit For
    End If
Next

Select Case clientnumbers
    Case Is = 1
        Set sh_paste = wsclient.Shapes("dwight")
    Case Is = 2
        Set sh_paste = wsclient.Shapes("jim")
    Case Is = 3
        Set sh_paste = wsclient.Shapes("mike")
    Case Is = 4
        Set sh_paste = wsclient.Shapes("stanley")
    Case Is = 5
        Set sh_paste = wsclient.Shapes("pam")
End Select

sh_paste.Copy
ws.Paste
Selection.Name = "CurrentClient"

Set sh_current = ws.Shapes("CurrentClient")

sh_current.Top = 258
sh_current.Left = 584

ws.Range("A1").Select
Application.CutCopyMode = False


End Sub

Public Sub ClearClients()

Dim sh As Shape

Dim ws As Worksheet
Dim wsclient As Worksheet

Set ws = Worksheets("DunderMifflin")
Set wsclient = Worksheets("clients")

For Each sh In ws.Shapes
    If sh.Name = "CurrentClient" Then
        sh.Delete
        Exit For
    End If
Next



End Sub

