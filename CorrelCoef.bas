Attribute VB_Name = "CorrelCoef"
Function getRanks(x As Variant) As Variant
    Dim hlpField() As Double
    Dim i As Integer
    Dim indicator As Boolean
    Dim swp As Variant
    Dim res() As Variant
    
    ReDim hlpField(UBound(x), 1)
    indicator = True
    
    'loading data to auxiliary array, adding index of an observation
    For i = 0 To UBound(x)
        hlpField(i, 0) = x(i)
        hlpField(i, 1) = i
    Next i
    
    'sorting according to magnitude of observed variable (bubble sort)
    While indicator
        indicator = False
        
        For i = 1 To UBound(x)
            If hlpField(i - 1, 0) > hlpField(i, 0) Then 'sorting according to observed value
                indicator = True
                'observed value
                swp = hlpField(i - 1, 0)
                hlpField(i - 1, 0) = hlpField(i, 0)
                hlpField(i, 0) = swp
                'index of observation must remain the same
                swp = hlpField(i - 1, 1)
                hlpField(i - 1, 1) = hlpField(i, 1)
                hlpField(i, 1) = swp
            End If
        Next i
    Wend
    
    'a rank is added to each observation instead of observed value (i.e. the minimum has rank 1, maximum n, n = number of observation)
    For i = 0 To UBound(x)
        hlpField(i, 0) = i + 1
    Next i
    
    indicator = True
    
    'sorting according to index of observation -> each observation is assigned with the rank of observed value
    While indicator
        indicator = False
        For i = 1 To UBound(x)
            If hlpField(i - 1, 1) > hlpField(i, 1) Then 'cislo pozorovani
                indicator = True
                
                swp = hlpField(i - 1, 0)
                hlpField(i - 1, 0) = hlpField(i, 0)
                hlpField(i, 0) = swp
                
                swp = hlpField(i - 1, 1)
                hlpField(i - 1, 1) = hlpField(i, 1)
                hlpField(i, 1) = swp
            End If
        Next i
    Wend
    
    ReDim res(UBound(x))
    
    'resulting array contains only ranks instead of , index of the observation is determined by the array index
    For i = 0 To UBound(x)
        res(i) = hlpField(i, 0)
    Next i
    
    getRanks = res
End Function

Function spearmanCorrel(x As Range, y As Range)
    Dim n As Integer
    Dim i As Integer
    Dim dataArrayX() As Variant
    Dim dataArrayY() As Variant
    
    If x.Rows.Count = y.Rows.Count And x.Columns.Count = 1 And y.Columns.Count = 1 Then
        n = 0
        'determing number of numeric values in input arrays x and y
        For i = 1 To x.Rows.Count
            If x.Cells(i, 1).Value <> "" And y.Cells(i, 1).Value <> "" _
               And IsNumeric(x.Cells(i, 1).Value) And IsNumeric(y.Cells(i, 1).Value) Then n = n + 1
        Next i
        
        'loading numeric values in input arrays to auxiliary arrays
        ReDim dataArrayX(n - 1)
        ReDim dataArrayY(n - 1)
        
        n = 0
        For i = 1 To x.Rows.Count
            If x.Cells(i, 1).Value <> "" And y.Cells(i, 1).Value <> "" _
               And IsNumeric(x.Cells(i, 1).Value) And IsNumeric(y.Cells(i, 1).Value) Then
                dataArrayX(n) = x.Cells(i, 1).Value
                dataArrayY(n) = y.Cells(i, 1).Value
                n = n + 1
            End If
        Next i
        
        'sorting data in arrays, dataArray X and Y now contain ranks of values of variables x and y for each observation
        dataArrayX = getRanks(dataArrayX)
        dataArrayY = getRanks(dataArrayY)
        
        'spearman coefficient calculation - s = 1 - 6sum(d_i^2)[(n(n^2-1)]
        For i = 0 To n - 1
            spearmanCorrel = spearmanCorrel + (dataArrayX(i) - dataArrayY(i)) ^ 2
        Next i
        
        spearmanCorrel = 1 - 6 * spearmanCorrel / (n * (n ^ 2 - 1))
    Else
        spearmanCorrel = xlErrNA
    End If
End Function

Function kendallCorrel(x As Range, y As Range)
    Dim n As Integer
    Dim i As Integer, j As Integer
    Dim c As Integer, d As Integer
    Dim dataArrayX() As Variant
    Dim dataArrayY() As Variant
    
    If x.Rows.Count = y.Rows.Count And x.Columns.Count = 1 And y.Columns.Count = 1 Then
        n = 0
        
        'determing number of numeric values in input arrays x and y
        For i = 1 To x.Rows.Count
            If x.Cells(i, 1).Value <> "" And y.Cells(i, 1).Value <> "" _
               And IsNumeric(x.Cells(i, 1).Value) And IsNumeric(y.Cells(i, 1).Value) Then n = n + 1
        Next i
        
        'loading numeric values in input arrays to auxiliary arrays
        ReDim dataArrayX(n - 1)
        ReDim dataArrayY(n - 1)
        
        n = 0
        For i = 1 To x.Rows.Count
            If x.Cells(i, 1).Value <> "" And y.Cells(i, 1).Value <> "" _
               And IsNumeric(x.Cells(i, 1).Value) And IsNumeric(y.Cells(i, 1).Value) Then
                dataArrayX(n) = x.Cells(i, 1).Value
                dataArrayY(n) = y.Cells(i, 1).Value
                n = n + 1
            End If
        Next i
        
        For i = 0 To n - 1
            'i - pivot observation
            For j = i + 1 To n - 1
                'checking all observation with index higher than pivot whether the direction of change is same in both x and y variables
                If dataArrayX(j) - dataArrayX(i) = 0 Then 'same rank of variable x as pivot, do nothing
                ElseIf dataArrayY(j) - dataArrayY(i) = 0 Then 'constant function (y same as in case of pivot)
                    c = c + 0.5
                    d = d + 0.5
                ElseIf Sgn(dataArrayY(j) - dataArrayY(i)) = Sgn(dataArrayX(j) - dataArrayX(i)) Then
                    'same direction of change => concordation, i.e. positive correlation
                    c = c + 1
                ElseIf Sgn(dataArrayY(j) - dataArrayY(i)) = -Sgn(dataArrayX(j) - dataArrayX(i)) Then
                    'oposite direction of change => discordation, i.e. negative correlation
                    d = d + 1
                End If
            Next j
        Next i
        
        kendallCorrel = (c - d) / (c + d) 'c+d - total number of comparisons, c-d - if c>d then correlation is positive
    Else
        kendallCorrel = xlErrNA
    End If
End Function


