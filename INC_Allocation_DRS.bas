Attribute VB_Name = "INC_Allocation_DRS"
Function YesGreen()
Attribute YesGreen.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.CutCopyMode = False
    Columns("C:C").FormatConditions.Add Type:=xlTextString, String:="Yes", _
        TextOperator:=xlContains
    Columns("C:C").FormatConditions(Columns("C:C").FormatConditions.Count).SetFirstPriority
    With Columns("C:C").FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Columns("C:C").FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Columns("C:C").FormatConditions(1).StopIfTrue = False

End Function
Function YesYellow()
Attribute YesYellow.VB_ProcData.VB_Invoke_Func = " \n14"

    Columns("D:D").FormatConditions.Add Type:=xlTextString, String:="Yes", _
        TextOperator:=xlContains
    Columns("D:D").FormatConditions(Columns("D:D").FormatConditions.Count).SetFirstPriority
    With Columns("D:D").FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Columns("D:D").FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Columns("D:D").FormatConditions(1).StopIfTrue = False
End Function
Function YesRed()
Attribute YesRed.VB_ProcData.VB_Invoke_Func = " \n14"

    Columns("B:B").FormatConditions.Add Type:=xlTextString, String:="Yes", _
        TextOperator:=xlContains
    Columns("B:B").FormatConditions(Columns("B:B").FormatConditions.Count).SetFirstPriority
    With Columns("B:B").FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Columns("B:B").FormatConditions(1).StopIfTrue = False
End Function
Function TextRed()
Attribute TextRed.VB_ProcData.VB_Invoke_Func = " \n14"

    With Range("J2:J10000").Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    
    With Range("L2:L10000").Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    
    With Range("N2:N10000").Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    
End Function

Function FillFormulaDown()
    ' Write the formula in cell B2
    Range("B2").Formula = "=IF(OR(P2=""Yes"", Q2=""Yes"", R2=""Yes""), ""Yes"", ""No"")"
    
    ' Move one cell to the right
    Range("B2").Offset(0, 1).Select
    
    ' Press Ctrl+Down
    Selection.End(xlDown).Select
    
    ' Move one cell to the left
    Selection.Offset(0, -1).Select
    
    ' Press Ctrl+Shift+Up
    Range(Selection, Selection.End(xlUp)).Select
    
    ' Fill the formula down
    Selection.FillDown
    
    Range("A1").Select
End Function

Function HeaderFormat()
    Dim cellRange1 As Range
    Dim cellRange2 As Range
    
    ' Define ranges
    Set cellRange1 = Range("A1,B1,C1,D1,E1,F1,G1,H1,I1,K1,M1,O1,P1,Q1,R1,S1,T1,U1")
    Set cellRange2 = Range("J1,L1,N1")
    
    ' Apply interior formatting to cellRange1
    With cellRange1.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Apply font formatting to cellRange1
    With cellRange1.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    ' Apply alignment formatting to cell F1
    With Range("F1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Apply interior formatting to cellRange2
    With cellRange2.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' Apply font formatting to cellRange2
    With cellRange2.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Function

Sub FilterAndCopyData()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim lastRow As Long
    Dim filterRange As Range
    Dim destHeaders As Variant
    Dim headerMap As Object
    Dim i As Long, j As Long, k As Long
    Dim header As String
    Dim srcCol As Range
    Dim destCol As Range
    
    ' Set the source worksheet
    Set wsSource = Worksheets("Sheet1")
    
    ' Find the last row with data in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range to filter (assuming the headers are in row 1 and data starts from row 2)
    Set filterRange = wsSource.Range("A1:AA" & lastRow) ' Adjust column AA to the last column of your data
    
    ' Apply filter on "I&C Status" column which is assumed to be column W (23rd column), adjust if different
    filterRange.AutoFilter Field:=23, Criteria1:="Yes"
    
    ' Check if the "Allocation" sheet already exists, if yes, delete it
    On Error Resume Next
    Set wsDestination = Worksheets("Allocation")
    If Not wsDestination Is Nothing Then
        Application.DisplayAlerts = False
        wsDestination.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create a new "Allocation" sheet
    Set wsDestination = Worksheets.Add
    wsDestination.Name = "Allocation"
    
    ' Define the desired headers in the specified order
    destHeaders = Array("Users", "Debit Interest", "Charges", "Credit Interest", "Amount", "Hit Date", "Sort Code", _
                        "Account", "Brand", "Accrued Amount", "Accrued Interest Rate", "Cutoff Amount", _
                        "Cutoff Interest Rate", "Applied Interest Amount", "Max Credit Interest Rate", _
                        "Accrued Interest", "Cutoff Interest", "Applied Interest", "Diarised Date", _
                        "Diary Amount", "Status")
    
    ' Create a dictionary to map source headers to destination headers
    Set headerMap = CreateObject("Scripting.Dictionary")
    
    ' Map the headers from the source sheet
    For i = 1 To wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        headerMap(wsSource.Cells(1, i).Value) = i
    Next i
    
    ' Find a template header cell for formatting
    Dim templateHeader As Range
    For i = 1 To wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
        If Not wsSource.Cells(1, i).Value = "" Then
            Set templateHeader = wsSource.Cells(1, i)
            Exit For
        End If
    Next i
    
    ' Write the headers to the destination sheet in the specified order
    For i = LBound(destHeaders) To UBound(destHeaders)
        header = destHeaders(i)
        ' Set the header value
        wsDestination.Cells(1, i + 1).Value = header
        ' Copy the format of the template header cell if the header is new
        If headerMap.Exists(header) Then
            wsSource.Cells(1, headerMap(header)).Copy
            wsDestination.Cells(1, i + 1).PasteSpecial Paste:=xlPasteFormats
        Else
            templateHeader.Copy
            wsDestination.Cells(1, i + 1).PasteSpecial Paste:=xlPasteFormats
        End If
        ' Reset the header value to ensure it doesn't change
        wsDestination.Cells(1, i + 1).Value = header
    Next i
    
    ' Copy the filtered data to the new sheet in the specified order, preserving number formats
    j = 2 ' Row counter for the destination sheet
    
    For i = 2 To lastRow
        If Not wsSource.Rows(i).Hidden Then
            For k = LBound(destHeaders) To UBound(destHeaders)
                header = destHeaders(k)
                If headerMap.Exists(header) Then
                    Set srcCol = wsSource.Cells(i, headerMap(header))
                    Set destCol = wsDestination.Cells(j, k + 1)
                    destCol.NumberFormat = srcCol.NumberFormat
                    destCol.Value = srcCol.Text
                End If
            Next k
            j = j + 1
        End If
    Next i
    
    ' Turn off the filter
    wsSource.AutoFilterMode = False
End Sub
Function SortAccount()

    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Allocation").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Allocation").AutoFilter.Sort.SortFields.Add2 Key:= _
        Range("H1:H671"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Allocation").AutoFilter.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    
    
End Function


Sub DRS_Output_Allocation()

Call FilterAndCopyData
Call HeaderFormat
Call FillFormulaDown
Call YesRed
Call YesGreen
Call YesYellow
Call TextRed
Call SortAccount
Call ColorDuplicatesInSelection

End Sub
