Sub homework2():

    ' Set Dimensions
    Dim total as Double

    ' Add Labels
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    ' Find the last row of data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount

        ' If ticker is not equal then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores result
            total = total + Cells(i, 7).Value

            ' Print ticker
            Range("I" & 2 + j).Value = Cells(i, 1).Value

            ' Print total
            Range("J" & 2 + j).Value = total

            ' Reset "total"
            total = 0

            ' Jump to next row
            j = j + 1

        ' Otherwise keep adding to "total"
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

End Sub