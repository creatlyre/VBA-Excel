Attribute VB_Name = "VisionCRR"
Option Explicit
Public fontRedInstruction As Boolean, currentVendor As String, VersantHealth As String, NumberPlan As Integer

Sub VisionCRR_Excel_USE_ME()
Dim ws As Worksheet, wsName() As String

currentVendor = InputBox("Provide a current incumbent:")
VersantHealth = InputBox("Who quoted as Versant Health?")

For Each ws In ThisWorkbook.Worksheets 'read through all worksheets
ws.Activate
    If InStr(1, LCase(Cells(1, 1)), "plan design summary") Then
        wsName = Split(Cells(1, 1), ":")
        ws.Name = Trim(wsName(1)) & " " & NumberPlan
        NumberPlan = NumberPlan + 1
        'Call the function whih works on the Vision Design Only
        Design
        'Debug.Print ws.Name
    Else
        'do notning at this point
    End If
    
Next

If fontRedInstruction = True Then
    MsgBox ("At least one vendor input Copay range without range just with one amount." & Chr(10) & Chr(10) & "Check all cells with red font!")
End If

End Sub

Sub Design()
Cells(1, 1).UnMerge
Cells(1, 1).Copy Cells(1, 3)
Columns(5).Delete
Columns(2).Delete
Columns(1).Delete
Range("A:A").ColumnWidth = 19
Range("B:Q").ColumnWidth = 10

Call ChangeVendorName("PlanDesign")

Dim RowsBenefits As Integer
Dim ColumnsBenefits As Integer, Frequencies() As String, AmountDivided() As String
'for loop changing order of benefits Copay: $10 --> $10 Copay etc.
For RowsBenefits = 4 To 500
    If IsEmpty(Cells(RowsBenefits, 1)) And IsEmpty(Cells(RowsBenefits + 1, 1)) Then
        Exit For
    Else
        On Error Resume Next
        For ColumnsBenefits = 2 To 14
            If IsEmpty(Cells(RowsBenefits, ColumnsBenefits)) Then
                GoTo NextValue
            ElseIf InStr(1, Cells(RowsBenefits, ColumnsBenefits), "/") Then 'correcting format of frequencies
                Frequencies = Split(Cells(RowsBenefits, ColumnsBenefits), "/")
                Cells(RowsBenefits, ColumnsBenefits).NumberFormat = "@"
                Cells(RowsBenefits, ColumnsBenefits) = Trim(Frequencies(0)) + "/" + Trim(Frequencies(1)) + "/" + Trim(Frequencies(2))
            'COPAY
            ElseIf InStr(1, LCase(Cells(RowsBenefits, ColumnsBenefits)), "copay") Then
                If Len(Cells(RowsBenefits, ColumnsBenefits)) < 14 Then
                    AmountDivided = Split(Cells(RowsBenefits, ColumnsBenefits))
                    Cells(RowsBenefits, ColumnsBenefits) = AmountDivided(1) & Chr(10) & Left(AmountDivided(0), Len(AmountDivided(0)) - 1)
                    Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
                'COPAY RANGE
                ElseIf InStr(1, LCase(Cells(RowsBenefits, ColumnsBenefits)), "range") And Len(Cells(RowsBenefits, ColumnsBenefits)) < 35 Then
                    AmountDivided = Split(Cells(RowsBenefits, ColumnsBenefits))
                    Dim WhereIsTo As Integer
                        For WhereIsTo = 0 To 10
                            If AmountDivided(WhereIsTo) = "" Then
                                Cells(RowsBenefits, ColumnsBenefits) = AmountDivided(WhereIsTo - 1) & Chr(10) & "Copay"
                                Cells(RowsBenefits, ColumnsBenefits).Font.Color = vbRed
                                Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
                                fontRedInstruction = True
                                Exit For
                            ElseIf AmountDivided(WhereIsTo) Like "*to*" Then
                                Cells(RowsBenefits, ColumnsBenefits) = AmountDivided(WhereIsTo - 1) & " - " & AmountDivided(WhereIsTo + 1) & Chr(10) & "Copay"
                                Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
                                Exit For
                            End If
                            
                        Next WhereIsTo
                        
                    
                End If
            'ALLOWANCE
            ElseIf InStr(1, LCase(Cells(RowsBenefits, ColumnsBenefits)), "allowance") And Len(Cells(RowsBenefits, ColumnsBenefits)) < 17 Then
                On Error GoTo ErrorHandler
                AmountDivided = Split(Cells(RowsBenefits, ColumnsBenefits))
                Cells(RowsBenefits, ColumnsBenefits) = AmountDivided(1) & Chr(10) & Left(AmountDivided(0), Len(AmountDivided(0)) - 1)
                Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
            'REIMBURSMENT
            ElseIf InStr(1, LCase(Cells(RowsBenefits, ColumnsBenefits)), "reimbursement") Then
                On Error GoTo ErrorHandler
                AmountDivided = Split(Cells(RowsBenefits, ColumnsBenefits))
                Cells(RowsBenefits, ColumnsBenefits) = AmountDivided(2) & Chr(10) & AmountDivided(0)
                Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
            'NOT COVERED
            ElseIf InStr(1, LCase(Cells(RowsBenefits, ColumnsBenefits)), "not covered") Then
                Cells(RowsBenefits, ColumnsBenefits) = "Not" & Chr(10) & "Covered"
                Cells(RowsBenefits, ColumnsBenefits).VerticalAlignment = xlVAlignCenter
                
ErrorHandler:
            If Err.Number = 9 Then
                MsgBox ("Subscript out of range" & Chr(10) & "Press Ok to continue")
            End If
            End If
            
NextValue:
        Next ColumnsBenefits
        
    End If
Next RowsBenefits
End Sub

Sub ChangeVendorName(WhichSlide As String)
Dim VendorName_Row As Integer, VendorName_Column As Integer, currentTrigger As Boolean
'currentVendor = "Ameritas"
'VersantHealth = "Superior Vision"

Select Case WhichSlide
       
Case "PlanDesign"
For VendorName_Row = 1 To 4
    For VendorName_Column = 1 To 20


        'Current is Ameritas
        If IsEmpty(Cells(VendorName_Row, VendorName_Column)) And IsEmpty(Cells(VendorName_Row, VendorName_Column + 2)) And IsEmpty(Cells(VendorName_Row, VendorName_Column + 3)) Then
            Exit For
        'HERE I WILL NEED SWITCH CASE SO I CAN DIFFER A DEPENDS ON WHICH SLIDE -- BELOW IS FOR PLAN DESIGN ONLY
        ElseIf InStr(1, LCase(Cells(VendorName_Row, VendorName_Column)), "ameritas") Then
            If currentVendor Like "*meritas*" And currentTrigger = False Then
                Cells(VendorName_Row, VendorName_Column) = "Ameritas" & Chr(10) & "Current"
                currentTrigger = True
            ElseIf currentVendor Like "*meritas*" And currentTrigger = True Then
                Cells(VendorName_Row, VendorName_Column) = "Ameritas" & Chr(10) & "Renewal"
                Exit Sub
            Else
                Cells(VendorName_Row, VendorName_Column) = "Ameritas"
                Exit Sub
            End If
            
        'Current is EyeMed
        ElseIf InStr(1, LCase(Cells(VendorName_Row, VendorName_Column)), "eyemed") And currentVendor Like "*ye?ed*" Then
            If currentTrigger = False Then
                Cells(VendorName_Row, VendorName_Column) = "EyeMed" & Chr(10) & "Current"
                currentTrigger = True
            ElseIf currentTrigger = True Then
                Cells(VendorName_Row, VendorName_Column) = "EyeMed" & Chr(10) & "Renewal"
                Exit Sub
            End If
            
        'Current is MetLife
        ElseIf InStr(1, LCase(Cells(VendorName_Row, VendorName_Column)), "eyemed") Then
            If currentVendor Like "*et?ife*" And currentTrigger = False Then
                Cells(VendorName_Row, VendorName_Column) = "MetLife" & Chr(10) & "Current"
                currentTrigger = True
            ElseIf currentVendor Like "*et?ife*" And currentTrigger = True Then
                Cells(VendorName_Row, VendorName_Column) = "MetLife" & Chr(10) & "Renewal"
                Exit Sub
            End If
            
        'Current is UHC
        ElseIf InStr(1, LCase(Cells(VendorName_Row, VendorName_Column)), "uhc") Then
            If currentVendor Like "*uhc*" Or currentVendor Like "*UHC*" Or currentVendor Like "?hc*" Then
                If currentTrigger = False Then
                    Cells(VendorName_Row, VendorName_Column) = "UHC" & Chr(10) & "Current"
                    currentTrigger = True
                ElseIf currentTrigger = True Then
                    Cells(VendorName_Row, VendorName_Column) = "UHC" & Chr(10) & "Renewal"
                    Exit Sub
                End If
            End If
            
        'Current is VersantHealth
        ElseIf InStr(1, LCase(Cells(VendorName_Row, VendorName_Column)), "versant") Then
            If currentVendor Like "*uperior*" Then
                If currentTrigger = False Then
                    Cells(VendorName_Row, VendorName_Column) = "Superior Vision" & Chr(10) & "Current"
                    currentTrigger = True
                ElseIf currentTrigger = True Then
                    Cells(VendorName_Row, VendorName_Column) = "Superior Vision" & Chr(10) & "Renewal"
                    Exit Sub
                End If
                
            ElseIf currentVendor Like "*avis*" Then
                If currentTrigger = False Then
                    Cells(VendorName_Row, VendorName_Column) = "Davis Vision" & Chr(10) & "Current"
                    currentTrigger = True
                ElseIf currentTrigger = True Then
                    Cells(VendorName_Row, VendorName_Column) = "Superior Vision" & Chr(10) & "Renewal"
                    Exit Sub
                End If
            Else
                If VersantHealth Like "*uperior*" Then
                    Cells(VendorName_Row, VendorName_Column) = "Superior Vision"
                ElseIf VersantHealth Like "*avis*" Then
                    Cells(VendorName_Row, VendorName_Column) = "Davis Vision"
                Else
                    MsgBox ("Can't find Versant Health" & Chr(10) & "Please change it manually.")
                End If
            End If
        End If



    Next VendorName_Column
Next VendorName_Row

Case Else
    
    'do nothing
End Select
End Sub
