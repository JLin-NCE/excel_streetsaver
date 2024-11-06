Sub CreateNewSortedSheet()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim currentCategory As String
    Dim lengthSum As Double
    Dim areaSum As Double
    Dim i As Long
    
    ' Reference the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row in the active sheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
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
    newWs.Range("P1").Value = "PCI Climate %"
    newWs.Range("Q1").Value = "PCI Other %"
    
    ' Format headers
    With newWs.Range("A1:Q1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Name = "Aptos Narrow"
        .Interior.Color = RGB(21, 61, 100)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 41
    End With
    
    ' Insert and format blank row after headers
    newWs.Rows("2:2").Insert
    With newWs.Rows("2:2")
        .RowHeight = 25
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
    End With
    newWs.Range("B2:C2").Merge
    
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
            newWs.Cells(cell.Row + 1, 7).Value = Mid(cell.Value, dashPos + 1)
        Else
            newWs.Cells(cell.Row + 1, 7).Value = cell.Value
        End If
    Next cell
    
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
    
    ' Set all data text to black
    newWs.Range("A3:Q" & (lastRow + 1)).Font.Color = vbBlack
    
    ' Initialize variables for category tracking and summing
    lastRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
    lengthSum = 0
    areaSum = 0
    currentCategory = newWs.Cells(3, 7).Value
    
    ' Add first category to first blank row
    newWs.Range("B2").Value = currentCategory
    
    ' Loop through rows to sum and insert summary rows for each category
    For i = 3 To lastRow + 1
        ' Check if we're at a category change or the end of data
        If i > lastRow Or (i <= lastRow And newWs.Cells(i, 7).Value <> currentCategory) Then
            ' Insert summary row when category changes or end of data
            newWs.Rows(i).Insert
            
            ' Format only the summary row cells that contain values
            With newWs.Rows(i)
                newWs.Cells(i, 8).Value = lengthSum / 5280 ' Convert length to miles
                newWs.Cells(i, 8).NumberFormat = "0.00"
                newWs.Cells(i, 8).Font.Bold = True
                
                newWs.Cells(i, 10).Value = areaSum
                newWs.Cells(i, 10).NumberFormat = "#,##0"
                newWs.Cells(i, 10).Font.Bold = True
                
                .Borders.LineStyle = xlContinuous
            End With
            
            ' Insert and format blank row after summary (only if not at the very end)
            If i <= lastRow Then
                newWs.Rows(i + 1).Insert
                With newWs.Rows(i + 1)
                    .RowHeight = 25
                    .Interior.Color = vbWhite
                    .Font.Color = vbBlack
                End With
                newWs.Range("B" & (i + 1) & ":C" & (i + 1)).Merge
                
                ' Get next category (if not at end) and add to blank row
                If i + 2 <= lastRow Then
                    newWs.Range("B" & (i + 1)).Value = newWs.Cells(i + 2, 7).Value
                End If
            End If
            
            ' Reset sums and update current category
            lengthSum = 0
            areaSum = 0
            If i + 2 <= lastRow Then
                currentCategory = newWs.Cells(i + 2, 7).Value
                lastRow = lastRow + 2 ' Account for summary and blank rows
                i = i + 1 ' Skip the blank row in next iteration
            End If
        End If
        
        ' Accumulate Length and Area sums for the current category
        If i <= lastRow Then
            If IsNumeric(newWs.Cells(i, 8).Value) Then lengthSum = lengthSum + CDbl(newWs.Cells(i, 8).Value)
            If IsNumeric(newWs.Cells(i, 10).Value) Then areaSum = areaSum + CDbl(newWs.Cells(i, 10).Value)
        End If
    Next i
    
    ' Final formatting
    With newWs.Range("A1:Q" & lastRow)
        .Borders.LineStyle = xlContinuous
    End With

    newWs.Columns("A:Q").AutoFit
    newWs.Activate
End Sub

