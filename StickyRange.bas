Attribute VB_Name = "StickyRange"
Function OneDayOverlap(r, c)

Dim h1, h0, l1, l0 As Single
Dim diff1, diff0 As Single
Dim u As Single
Dim sshLast, sshVwap, sshHigh, sshLow As Worksheet

    Set sshLast = Worksheets("Last")
    Set sshVwap = Worksheets("VWAP")
    Set sshHigh = Worksheets("High")
    Set sshLow = Worksheets("Low")
    
    h1 = sshHigh.Cells(r, c).Value
    h0 = sshHigh.Cells(r - 1, c).Value
    l1 = sshLow.Cells(r, c).Value
    l0 = sshLow.Cells(r - 1, c).Value
    
    diff1 = h1 - l1
    diff0 = h0 - l0
    
    If h0 >= h1 Then
        If l0 < l1 Then
            u = 1
        Else
            over = h1 - l0
            u = WorksheetFunction.Max(over / diff0, over / diff1)
        End If
    Else
        If l0 > l1 Then
            u = 1
        Else
            over = h0 - l1
            u = WorksheetFunction.Max(over / diff0, over / diff1)
        End If
    End If
    
    OneDayOverlap = u
                
End Function
