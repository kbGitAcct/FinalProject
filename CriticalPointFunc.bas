Attribute VB_Name = "CriticalPointFunc"
Option Base 1
Public Const SymbolRow As Integer = 3
Sub calcAllCPs()
    Set e = Sheets("Last").Cells(SymbolRow, 2)
    
    Sheets("CriticalPts").Range("B" & SymbolRow + 1 & ":CAA4000").ClearContents
    Do While Len(e) > 0
        n = e.Column
        Call calcCPs(n)
        Set e = e.Offset(0, 1)
    Loop
End Sub
Sub calcCPs(c)

Dim firstPass As Boolean
Dim r, slopeCount As Integer
Dim m As Single
Dim equity, firstPx As Object
Dim sshLast, sshSlope As Worksheet
        
    Set sshLast = Worksheets("Last")
    Set sshCP = Worksheets("CriticalPts")

    Set equity = sshLast.Range("B" & SymbolRow)
    slopeCount = sshCP.Range("A1").Value
            
    lastRow = sshLast.Cells(10000, c).End(xlUp).Row
        
    Set firstPx = sshLast.Cells(equity.Offset(1, 0).Row, c)
    initCount = 0
    Do While initCount < slopeCount And Len(firstPx) > 0
        If Left(firstPx, 1) <> "#" Then
            If Len(firstPx) > 0 Then initCount = initCount + 1
        Else
            initCount = 0
        End If
        
        Set firstPx = firstPx.Offset(1, 0)
    Loop
    
    If initCount < slopeCount Then Exit Sub
        
    mPrev = 0
    lastInflexRow = firstPx.Row - (slopeCount - 1)
    firstPass = True
    Do While firstPx.Row < lastRow
        r = firstPx.Row
        m = Round(getSlope(r - (slopeCount - 1), r, c), 4)
        
        If m * mPrev < 0 Then
            If m < 0 Then
                currInflex = WorksheetFunction.Max(Range(sshLast.Cells(lastInflexRow, c), sshLast.Cells(r, c)))
            Else
                currInflex = WorksheetFunction.Min(Range(sshLast.Cells(lastInflexRow, c), sshLast.Cells(r, c)))
            End If
            
            For i = r To lastInflexRow Step -1
                If sshLast.Cells(i, c) = currInflex Then
                    currInflexRow = i
                    Exit For
                End If
            Next i
            
            sshCP.Cells(r, c).Value = Sgn(m) * currInflexRow
            
            lastInflexRow = currInflexRow
            mPrev = m
        End If
        If firstPass = True Then
            mPrev = m
            firstPass = False
        End If
        Set firstPx = firstPx.Offset(1, 0)
    Loop
                
End Sub
Function getSlope(r1, r2, c)

Dim yArray, xArray As Variant

    yArray = setYarray(r1, r2, c)
    
    yN = UBound(yArray)
    
    If yN > (0.66 * (r2 - r1)) Then
        xArray = setXarray(yN)
        a = UBound(xArray)
        d = UBound(yArray)
        m = WorksheetFunction.Slope(yArray, xArray)
    Else
        m = 0
    End If
    
    getSlope = m
End Function
Function setXarray(n)

Dim nArray As Variant
    
    ReDim nArray(n)
    
    For i = 1 To n
        nArray(i) = i
    Next i
    
    setXarray = nArray
    
End Function
Function setYarray(a, b, col)

Dim nArray As Variant
Dim sshLast As Worksheet
        
    Set sshLast = Worksheets("Last")
    
    ReDim nArray(1)
    k = 1
    For i = a To b
        thisPx = sshLast.Cells(i, col).Value
        If Left(thisPx, 1) <> "#" And Len(thisPx) > 0 Then
            ReDim Preserve nArray(k)
            nArray(k) = CSng(sshLast.Cells(i, col))
            k = k + 1
        End If
    Next i
    
    setYarray = nArray
    
End Function

