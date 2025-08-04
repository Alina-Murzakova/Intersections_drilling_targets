Attribute VB_Name = "functions"

'Проверка ориентации

Function CheckOrientation(x1 As Variant, y1 As Variant, x2 As Variant, y2 As Variant, x3 As Variant, y3 As Variant)
    Dim Orientation As Double
    Dim result As Integer
    
    Orientation = (y2 - y1) * (x3 - x2) - (y3 - y2) * (x2 - x1)
    
    If x2 = Empty Or x3 = Empty Then
        result = 0
    ElseIf Orientation > 0 Then 'По часовой стрелке
        result = 1
    ElseIf Number < 0 Then 'Против часовой стрелки
        result = 2
    Else ' Коллинеарны
        result = 0
    End If
    
    CheckOrientation = result

End Function


'Проверка пересечений отрезков

Function Intersect(x1_current As Variant, y1_current As Variant, x3_current As Variant, y3_current As Variant, _
                    x1_round As Variant, y1_round As Variant, x3_round As Variant, y3_round As Variant)
    
    If CheckOrientation(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round) <> CheckOrientation(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round) And _
            CheckOrientation(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current) <> CheckOrientation(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current) Then
        Intersect = True
    End If
    
    
End Function

' Расстояние между двумя точками

Function LenBetweenPoints(x1, y1, x2, y2)

    LenBetweenPoints = ((x2 - x1) ^ 2 + (y2 - y1) ^ 2) ^ 0.5
    
End Function

' Минимальное расстояние между отрезком и точкой

Function LenBetweenSegmentPoint(x1, y1, x3, y3, x_round, y_round)
    Dim L1 As Double
    Dim L2 As Double
    Dim L As Double
    Dim x_base As Double
    Dim y_base As Double
    Dim A As Double
    Dim B As Double
    Dim C As Double
    Dim P As Double
    
    L1 = LenBetweenPoints(x_round, y_round, x1, y1) 'Расстояние от одного конца отрезка до конца скважины окружения
    L2 = LenBetweenPoints(x_round, y_round, x3, y3) 'Расстояние от другого конца отрезка до конца скважины окружения
    L = LenBetweenPoints(x1, y1, x3, y3) ' Длина ГС
      
    If (L1 * L1 > L2 * L2 + L * L) Or (L2 * L2 > L1 * L1 + L * L) Then
        P = WorksheetFunction.Min(L1, L2)
    Else
        If x1 = x3 And y1 <> y3 Then
            x_base = x1
            y_base = y_round
        ElseIf y1 = y3 And x1 <> x3 Then
            x_base = x_round
            y_base = y1
        ElseIf (x1 = x3 And y1 = y3) Or (IsEmpty(x3) And IsEmpty(y3)) Then
            x_base = x1
            y_base = y1
        Else
            A = y3 - y1
            B = x1 - x3
            C = -1 * x1 * (y3 - y1) + y1 * (x3 - x1)

            x_base = (B * x_round / A - C / B - y_round) * A * B / (A * A + B * B)
            y_base = B * x_base / A + y_round - B * x_round / A
            'Debug.Print (x_base)
            'Debug.Print (y_base)
        End If

        P = LenBetweenPoints(x_round, y_round, x_base, y_base)
    End If
    LenBetweenSegmentPoint = P
    

End Function


' Минимальное расстояние между двумя непересекающимися отрезками
Function MinLength(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round, x3_round, y3_round)
    
    MinLength = WorksheetFunction.Min(LenBetweenSegmentPoint(x1_current, y1_current, x3_current, y3_current, x1_round, y1_round), _
                                    LenBetweenSegmentPoint(x1_current, y1_current, x3_current, y3_current, x3_round, y3_round), _
                                LenBetweenSegmentPoint(x1_round, y1_round, x3_round, y3_round, x1_current, y1_current), _
                            LenBetweenSegmentPoint(x1_round, y1_round, x3_round, y3_round, x3_current, y3_current))
    
End Function

' Преобразование коллеции в массив
Function CollectionToArray(myCol As Collection) As Variant

    Dim result As Variant
    Dim cnt As Long

    ReDim result(myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result

End Function
   
Sub main()
    Dim res As Boolean
    Dim length As Double

    'res = Intersect(585303, 6963317, 584918, 6962857, 585187, 6963093, 584614, 6962692)
    'Debug.Print (res)
    'length = LenBetweenPoints(1, 0, 5, 0)
    'Debug.Print (length)
    'leng = LenBetweenSegmentPoint(585303, 6963317, 584918, 6962857, 585187, 6963093)
    'Debug.Print (leng)
    minlen = MinLength(587367, 6963242, 587753, 6962782, 587332, 6963289, 587572, 6962631)
    'Debug.Print (minlen)
    
End Sub
