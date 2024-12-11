Attribute VB_Name = "Module1"
' Version 1.0 - release, highlights & unhilights
' Written by Zac Spitzer, October 23 2024
' Last updated December 11 2024

Sub CopyCheck()
Application.ScreenUpdating = False
CheckPosition
CheckDepartment
CheckLocation
Application.ScreenUpdating = True
End Sub

Sub CheckPosition()

Dim Wksht As Worksheet
Dim Validation As Worksheet
Dim lastrow As Long
Dim Starte As Range
Dim Ende As Range
Dim ScanRange As Range

Set Wksht = ActiveSheet
Set Validation = Sheets("Data Validation")
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Set Starte = Wksht.Range("E5")
Set Ende = Wksht.Range("E" & lastrow)
Set ScanRange = Wksht.Range(Starte, Ende)

For Each Cell In ScanRange
If Cell.Value = Validation.Range("A2") Or _
   Cell.Value = Validation.Range("A3") Or _
   Cell.Value = Validation.Range("A4") Or _
   Cell.Value = Validation.Range("A5") Or _
   Cell.Value = Validation.Range("A6") Or _
   Cell.Value = Validation.Range("A7") Or _
   Cell.Value = Validation.Range("A8") Or _
   Cell.Value = Validation.Range("A9") Or _
   Cell.Value = Validation.Range("A10") Or _
    Cell.Value = Validation.Range("A11") Or _
    Cell.Value = Empty Then
        Cell.Select
        ActiveCell.Interior.ColorIndex = 0
    Else
        Cell.Select
        ActiveCell.Interior.ColorIndex = 6
End If
Next

End Sub

Sub CheckDepartment()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim Startf As Range
Dim Endf As Range
Dim ScanRange As Range

Set Wksht = ActiveSheet
Set Validation = Sheets("Data Validation")
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Set Startf = Wksht.Range("F5")
Set Endf = Wksht.Range("F" & lastrow)
Set ScanRange = Wksht.Range(Startf, Endf)

For Each Cell In ScanRange
If Cell.Value = Validation.Range("C2") Or _
   Cell.Value = Validation.Range("C3") Or _
   Cell.Value = Validation.Range("C4") Or _
   Cell.Value = Validation.Range("C5") Or _
   Cell.Value = Validation.Range("C6") Or _
   Cell.Value = Validation.Range("C7") Or _
   Cell.Value = Validation.Range("C8") Or _
   Cell.Value = Validation.Range("C9") Or _
   Cell.Value = Validation.Range("C10") Or _
    Cell.Value = Validation.Range("C11") Or _
    Cell.Value = Empty Then
        Cell.Select
        ActiveCell.Interior.ColorIndex = 0
    Else
        Cell.Select
        ActiveCell.Interior.ColorIndex = 6
End If
Next
End Sub


Sub CheckLocation()
Dim Wksht As Worksheet
Dim lastrow As Long
Dim Startt As Range
Dim Endt As Range
Dim ScanRange As Range

Set Wksht = ActiveSheet
Set Validation = Sheets("Data Validation")
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Set Startt = Wksht.Range("T5")
Set Endt = Wksht.Range("T" & lastrow)
Set ScanRange = Wksht.Range(Startt, Endt)

For Each Cell In ScanRange
If Cell.Value = Validation.Range("B2") Or _
   Cell.Value = Validation.Range("B3") Or _
   Cell.Value = Validation.Range("B4") Or _
   Cell.Value = Validation.Range("B5") Or _
   Cell.Value = Validation.Range("B6") Or _
   Cell.Value = Validation.Range("B7") Or _
   Cell.Value = Validation.Range("B8") Or _
   Cell.Value = Validation.Range("B9") Or _
   Cell.Value = Validation.Range("B10") Or _
    Cell.Value = Validation.Range("B11") Or _
    Cell.Value = Empty Then
        Cell.Select
        ActiveCell.Interior.ColorIndex = 0
    Else
        Cell.Select
        ActiveCell.Interior.ColorIndex = 6
End If
Next
End Sub

