Sub homework2():

    ' Set Dimensions
    Dim ws As Worksheet
    Dim total As Double
    Dim j As Integer

    For Each ws In Worksheets
    
        ' Set variables for each sheet
        total = 0
        j = 0

        ' Add Labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"

        ' Find the last row of data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker is not equal then print results
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores results 
                total = total + ws.Cells(i, 7).Value

                ' Print ticker
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value

                ' Print total
                ws.Range("J" & 2 + j).Value = total

                ' Reset "total"
                total = 0
                ' Jump to next row
                j = j + 1

            ' Otherwise keep adding to "total"
            Else
                total = total + ws.Cells(i, 7).Value

            End If

        Next i

        ' cleanup for next worksheet
        total = 0
        j = 0

    Next ws

End Sub
