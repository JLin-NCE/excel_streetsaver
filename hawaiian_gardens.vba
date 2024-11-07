Sub CreateNewSortedSheet()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim currentCategory As String
    Dim lengthSumRange As String
    Dim areaSumRange As String
    Dim i As Long
    Dim startRow As Long
    Dim categoryStartRow As Long
    Dim rowCount As Long
    
    ' Reference the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in the active sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Sort the active sheet by Functional Class
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("I2:I" & lastRow), _
            Order:=xlAscending, _
            CustomOrder:="Arterial,Collector,Residential/Local,Other"
        .SetRange ws.Range("A1:AJ" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    ' Delete existing summary sheets
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("PCI Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Add new worksheet
    Set newWs = ThisWorkbook.Worksheets.Add(After:=ws)
    newWs.Name = "PCI Report"
    
    ' Copy headers
    newWs.Range("A1").Value = "Street ID"
    newWs.Range("B1").Value = "Section ID"
    newWs.Range("C1").Value = "Street Name"
    newWs.Range("D1").Value = "From"
    newWs.Range("E1").Value = "To"
    newWs.Range("F1").Value = "Lanes"
    newWs.Range("G1").Value = "Functional Class"
    newWs.Range("H1").Value = "Length"
    newWs.Range("I1").Value = "Width"
    newWs.Range("J1").Value = "Area"
    newWs.Range("K1").Value = "Surface Type"
    newWs.Range("L1").Value = "Area ID"
    newWs.Range("M1").Value = "Insp. Date"
    newWs.Range("N1").Value = "PCI"
    newWs.Range("O1").Value = "PCI Load %"
    newWs.Range("Q1").Value = "PCI Climate %"
    newWs.Range("P1").Value = "PCI Other %"
    
    ' Copy data with new column mappings (adjusted for blank row)
    ws.Range("A2:A" & lastRow).Copy newWs.Range("A3") ' Street ID
    ws.Range("B2:B" & lastRow).Copy newWs.Range("B3") ' Section ID
    ws.Range("C2:C" & lastRow).Copy newWs.Range("C3") ' Street Name
    ws.Range("D2:D" & lastRow).Copy newWs.Range("D3") ' From
    ws.Range("E2:E" & lastRow).Copy newWs.Range("E3") ' To
    ws.Range("H2:H" & lastRow).Copy newWs.Range("F3") ' Lanes
    
    ' Copy and truncate Functional Class (Column G)
    Dim cell As Range
    Dim dashPos As Long
    For Each cell In ws.Range("I2:I" & lastRow)
        dashPos = InStr(cell.Value, "-")
        If dashPos > 0 Then
            newWs.Cells(cell.row + 1, 7).Value = Mid(cell.Value, dashPos + 1)
        Else
            newWs.Cells(cell.row + 1, 7).Value = cell.Value
        End If
    Next cell
    
    ' Copy row 3 column G value to row 2 and format row 2
    newWs.Range("B2").Value = newWs.Range("G3").Value
    With newWs.Range("A2:Q2")
        .Font.Bold = True
        .Font.Italic = True
        .Font.Size = 14
        .RowHeight = 25
    End With
    
    ' Ensure columns B and C are the same width before merging
    newWs.Columns("B:C").ColumnWidth = newWs.Columns("B").ColumnWidth
    ' Now merge the cells
    With newWs.Range("B2:C2")
        .MergeCells = True
        .HorizontalAlignment = xlLeft
    End With

ws.Range("J2:J" & lastRow).Copy newWs.Range("H3") ' Length
    ws.Range("K2:K" & lastRow).Copy newWs.Range("I3") ' Width
    ws.Range("L2:L" & lastRow).Copy newWs.Range("J3") ' Area
    ws.Range("Q2:Q" & lastRow).Copy newWs.Range("K3") ' Surface Type
    ws.Range("X2:X" & lastRow).Copy newWs.Range("L3") ' Area ID
    ws.Range("AD2:AD" & lastRow).Copy newWs.Range("M3") ' Insp. Date
    ws.Range("AB2:AB" & lastRow).Copy newWs.Range("N3") ' PCI
    ws.Range("AH2:AH" & lastRow).Copy newWs.Range("O3") ' PCI Load %
    ws.Range("AI2:AI" & lastRow).Copy newWs.Range("P3") ' PCI Climate %
    ws.Range("AJ2:AJ" & lastRow).Copy newWs.Range("Q3") ' PCI Other %
    
    ' Initialize variables for category tracking and summing
    lastRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).row
    currentCategory = newWs.Cells(3, 7).Value
    startRow = 3
    categoryStartRow = startRow
    rowCount = 0
    
    ' Create analysis sheet
    Dim analysisWs As Worksheet
    Set analysisWs = ThisWorkbook.Worksheets.Add(After:=newWs)
    analysisWs.Name = "Category Analysis"
    
    ' Set up analysis headers
    analysisWs.Range("A1").Value = "Category"
    analysisWs.Range("B1").Value = "Start Row"
    analysisWs.Range("C1").Value = "End Row"
    analysisWs.Range("D1").Value = "Row Count"
    analysisWs.Range("E1").Value = "Total Length (mi)"
    analysisWs.Range("F1").Value = "Total Area"
    
    Dim analysisRow As Long
    analysisRow = 2
    
    ' Loop through rows to sum and insert summary rows for each category
    For i = 3 To lastRow
        rowCount = rowCount + 1
        
        ' Check if we're at a category change OR at the last row
        If i = lastRow Or (newWs.Cells(i + 1, 7).Value <> currentCategory And newWs.Cells(i + 1, 7).Value <> "") Then
            ' Add to analysis sheet
            analysisWs.Cells(analysisRow, 1).Value = currentCategory
            analysisWs.Cells(analysisRow, 2).Value = categoryStartRow
            analysisWs.Cells(analysisRow, 3).Value = i
            analysisWs.Cells(analysisRow, 4).Value = rowCount
            analysisWs.Cells(analysisRow, 5).Formula = "=TEXT(ROUND(SUM(H" & categoryStartRow & ":H" & i & ")/5280,1),""0.0"")"
            analysisWs.Cells(analysisRow, 6).Formula = "=ROUND(SUM(J" & categoryStartRow & ":J" & i & "),1)"
            analysisRow = analysisRow + 1
            
            ' Insert summary row with formulas
            newWs.Rows(i + 1).Insert
            With newWs.Range("A" & (i + 1) & ":Q" & (i + 1))
                .Font.Bold = True
            End With
            With newWs.Range("A" & (i + 1) & ",D" & (i + 1) & ":Q" & (i + 1))
                .HorizontalAlignment = xlCenter
            End With
            newWs.Cells(i + 1, 8).Formula = "=TEXT(ROUND(SUM(H" & categoryStartRow & ":H" & i & ")/5280,1),""0.0"")"
            newWs.Cells(i + 1, 10).Formula = "=ROUND(SUM(J" & categoryStartRow & ":J" & i & "),1)"
            
            ' Insert blank row after summary if not at the end
            If i < lastRow Then
                newWs.Rows(i + 2).Insert
                newWs.Range("B" & (i + 2)).Value = newWs.Cells(i + 3, 7).Value
                With newWs.Range("A" & (i + 2) & ":Q" & (i + 2))
                    .Font.Bold = True
                    .Font.Italic = True
                    .Font.Size = 14
                    .RowHeight = 25
                End With
                newWs.Range("B" & (i + 2) & ":C" & (i + 2)).Merge
                
                ' Reset for the next category
                rowCount = 0
                currentCategory = newWs.Cells(i + 3, 7).Value
                categoryStartRow = i + 3
                lastRow = lastRow + 2
                i = i + 2
            End If
        End If
    Next i
' Check if a third category exists and add a final summary row with formulas if needed
    If analysisRow >= 4 Then
        Debug.Print "Third category starts 3 rows after previous and continues from row:", analysisWs.Cells(analysisRow - 1, 3).Value + 3, "to row:", lastRow
        
        ' Insert summary for third category with formulas
        newWs.Rows(lastRow + 1).Insert
        With newWs.Range("A" & (lastRow + 1) & ":Q" & (lastRow + 1))
            .Font.Bold = True
        End With
        With newWs.Range("A" & (lastRow + 1) & ",D" & (lastRow + 1) & ":Q" & (lastRow + 1))
            .HorizontalAlignment = xlCenter
        End With
        newWs.Cells(lastRow + 1, 8).Formula = "=TEXT(ROUND(SUM(H" & categoryStartRow & ":H" & lastRow & ")/5280,1),""0.0"")"
        newWs.Cells(lastRow + 1, 10).Formula = "=ROUND(SUM(J" & categoryStartRow & ":J" & lastRow & "),1)"
        
        ' Add final blank row and merge B:C
        newWs.Rows(lastRow + 2).Insert
        With newWs.Range("A" & (lastRow + 2) & ":Q" & (lastRow + 2))
            .Font.Bold = True
            .Font.Italic = True
            .Font.Size = 14
            .RowHeight = 25
        End With
        newWs.Range("B" & (lastRow + 2) & ":C" & (lastRow + 2)).Merge
    End If
    
    ' Format analysis sheet
    With analysisWs.Range("A1:F1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Name = "Aptos Narrow"
        .Interior.Color = RGB(21, 61, 100)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 41
    End With
    
    analysisWs.Columns("A:F").AutoFit
    
    ' Updated styling for PCI Report
    With newWs.Range("A1:Q1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Name = "Aptos Narrow"
        .Interior.Color = RGB(21, 61, 100)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 41
    End With

    ' Clear interior color for all cells except header row
    newWs.Range("A2:Q" & lastRow).Interior.ColorIndex = xlNone
    
    ' Make sure text is black for data rows
    With newWs.Range("A2:Q" & lastRow)
        .Font.Color = vbBlack
    End With

    ' Apply borders only to rows with content
    Dim rng As Range
    Dim rowHasContent As Boolean
    Dim row As Long
    
    ' Always apply borders to header row
    With newWs.Range("A1:Q1")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With
    
    ' Check each row for content and apply borders accordingly
    For row = 2 To lastRow
        rowHasContent = False
        Set rng = newWs.Range("A" & row & ":Q" & row)
        
        ' Check if row has any content
        For Each cell In rng
            If Not isEmpty(cell) Then
                rowHasContent = True
                Exit For
            End If
        Next cell
        
        ' Apply borders if row has content
        If rowHasContent Then
            With rng
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
            End With
        End If
    Next row
    
    ' Special formatting for category title rows
    Dim titleRow As Range
    For row = 2 To lastRow
        If newWs.Range("B" & row).MergeCells Then
            Set titleRow = newWs.Range("A" & row & ":Q" & row)
            With titleRow
                .Font.Bold = True
                .Font.Italic = True
                .Font.Size = 14
                .RowHeight = 25
                .Interior.ColorIndex = xlNone
            End With
        End If
    Next row
' Special formatting for summary rows (rows with formulas)
    For row = 2 To lastRow
        Set cell = newWs.Range("H" & row)
        If Left(cell.Formula, 1) = "=" Then
            With newWs.Range("A" & row & ":Q" & row)
                .Font.Bold = True
                .Interior.ColorIndex = xlNone
            End With
        End If
    Next row
    
    ' Autofit columns
    newWs.Columns("A:Q").AutoFit
    newWs.Activate
    
    ' Delete analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Category Analysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Now move the Other section
    MoveOtherSectionInNewSheet newWs
End Sub

Sub MoveOtherSectionInNewSheet(ws As Worksheet)
    ' Create collection to store sections
    Dim sections As Collection
    Set sections = New Collection
    
    ' Variables for tracking current section
    Dim sectionName As String
    Dim sectionStart As Long
    Dim sectionEnd As Long
    Dim i As Long
    Dim isInSection As Boolean
    isInSection = False
    
    ' Find the last row in the worksheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Process each row to identify sections
    For i = 1 To lastRow
        ' Check if this is a section header (empty column A, value in B, empty in C)
        If isEmpty(ws.Cells(i, 1)) And Not isEmpty(ws.Cells(i, 2)) And isEmpty(ws.Cells(i, 3)) Then
            ' If we were tracking a section, add it to our collection
            If isInSection Then
                ' Store section info as concatenated string: "name|startRow|endRow"
                sections.Add sectionName & "|" & sectionStart & "|" & (i - 1)
            End If
            
            ' Start new section
            sectionName = ws.Cells(i, 2).Value
            sectionStart = i + 1
            isInSection = True
        End If
    Next i
    
    ' Add the last section if we were tracking one
    If isInSection Then
        sections.Add sectionName & "|" & sectionStart & "|" & lastRow
    End If
    
    ' Find "Other" section
    Dim otherStartRow As Long
    Dim otherEndRow As Long
    Dim sectionInfo As Variant
    Dim sectionParts() As String
    
    otherStartRow = 0
    otherEndRow = 0
    
    For i = 1 To sections.Count
        sectionInfo = sections(i)
        sectionParts = Split(sectionInfo, "|")
        
        If InStr(1, sectionParts(0), "Other", vbTextCompare) > 0 Then
            otherStartRow = CLng(sectionParts(1)) - 1  ' Include header row
            otherEndRow = CLng(sectionParts(2))
            Exit For
        End If
    Next i
    
    ' If we found the Other section, move it
    If otherStartRow > 0 Then
        ' Copy the Other section including headers
        Dim otherRange As Range
        Set otherRange = ws.Range("A" & otherStartRow & ":Q" & otherEndRow)
        
        ' Store the data in an array
        Dim otherData As Variant
        otherData = otherRange.Value
        
        ' Clear the original range
        otherRange.Clear
        
        ' Find the new bottom of the worksheet
        Dim newLastRow As Long
        newLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 2  ' Add 2 for extra blank row
        
        ' Paste the Other section at the bottom
        ws.Range("A" & newLastRow).Resize(UBound(otherData, 1), UBound(otherData, 2)) = otherData
        
        ' Delete blank rows in the original range
        Dim rng As Range
        Dim cell As Range
        Dim isRowEmpty As Boolean
        
        For i = otherEndRow To otherStartRow Step -1
            Set rng = ws.Range("A" & i & ":Q" & i)
            isRowEmpty = True
            
            ' Check if row is empty
            For Each cell In rng
                If Not isEmpty(cell) Then
                    isRowEmpty = False
                    Exit For
                End If
            Next cell
            
            ' Delete if empty
            If isRowEmpty Then
                rng.EntireRow.Delete
            End If
        Next i
        
        ' Apply borders only to rows with content in the moved section
        For i = newLastRow To newLastRow + UBound(otherData, 1) - 1
            rowHasContent = False
            Set rng = ws.Range("A" & i & ":Q" & i)
            
            ' Check if row has any content
            For Each cell In rng
                If Not isEmpty(cell) Then
                    rowHasContent = True
                    Exit For
                End If
            Next cell
            
            ' Apply borders if row has content
            If rowHasContent Then
                With rng
                    .Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlInsideVertical).LineStyle = xlContinuous
                End With
            End If
        Next i
        
        ' Ensure proper text color in moved section
        ws.Range("A" & newLastRow & ":Q" & (newLastRow + UBound(otherData, 1) - 1)).Font.Color = vbBlack
    End If
    
    ' Final formatting cleanup
    ws.Activate
    ws.Columns("A:Q").AutoFit
    

End Sub

Sub CountRowsWithTextAndStyle()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim rowHasContent As Boolean
    Dim totalRows As Long
    Dim lastContentRow As Long
    Dim rowsWithContent As Collection
    Set rowsWithContent = New Collection
    
    ' Reference the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in the worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1  ' Add 1 to include extra row
    
    ' Initialize counter and last content row tracker
    totalRows = 0
    lastContentRow = 0
    
    ' First, collect all rows that have content
    For i = 2 To lastRow ' Start from 2 to skip title row
        rowHasContent = False
        Set rng = ws.Range("A" & i & ":Q" & i)
        
        ' Check each cell in the row for content
        For Each cell In rng
            If Not IsEmpty(cell) And Trim(CStr(cell.Text)) <> "" Then
                rowHasContent = True
                Exit For
            End If
        Next cell
        
        ' If row has content or it's the extra row after content, add to collection
        If rowHasContent Then
            totalRows = totalRows + 1
            lastContentRow = i
            rowsWithContent.Add i
        ElseIf i = lastContentRow + 1 Then
            ' Add the extra row after the last content row
            rowsWithContent.Add i
            lastContentRow = i  ' Update lastContentRow to include this extra row
        End If
    Next i
    
    ' Now apply styling to all content rows
    For Each rowNum In rowsWithContent
        Set rng = ws.Range("A" & rowNum & ":Q" & rowNum)
        
        ' Clear any existing background
        rng.Interior.ColorIndex = xlNone
        
        ' Apply borders
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        
        ' Special formatting for last row (which is now the extra row)
        If CLng(rowNum) = lastContentRow Then
            With rng
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next rowNum
    
    ' Display confirmation
    MsgBox "Styling applied to " & rowsWithContent.Count & " rows (including extra row)." & vbNewLine & _
           "Last styled row: " & lastContentRow, vbInformation, "Styling Complete"
End Sub
