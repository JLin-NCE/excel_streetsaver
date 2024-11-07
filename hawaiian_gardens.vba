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
            newWs.Cells(cell.Row + 1, 7).Value = Mid(cell.Value, dashPos + 1)
        Else
            newWs.Cells(cell.Row + 1, 7).Value = cell.Value
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
    newWs.Range("B2:C2").Merge
    
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
    lastRow = newWs.Cells(newWs.Rows.Count, "A").End(xlUp).Row
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
    
    ' Formatting applied at the end for PCI Report
    With newWs.Range("A1:Q1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Font.Name = "Aptos Narrow"
        .Interior.Color = RGB(21, 61, 100)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .RowHeight = 41
    End With

    With newWs.Range("A3:Q" & lastRow)
        .Font.Color = vbBlack
    End With

    ' Set borders for main range
    With newWs.Range("A1:Q" & lastRow)
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Special formatting for last row and the row after
    With newWs.Range("A" & lastRow & ":Q" & lastRow)
        .RowHeight = 25
    End With
    
    With newWs.Range("A" & (lastRow + 1) & ":Q" & (lastRow + 1))
        .Borders.LineStyle = xlContinuous
    End With
    
    newWs.Columns("A:Q").AutoFit
    newWs.Activate
    
    ' Delete analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Category Analysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub


