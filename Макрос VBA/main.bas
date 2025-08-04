Attribute VB_Name = "main"
Sub SearchIntersection()
'Количество строк на листе

Dim FieldUniques As New Collection
Dim FieldColumn As Range
Dim CellField As Range
Dim num_intersection As Integer
Dim ColumnStartHoleCurrent As New Collection
Dim ColumnStartHoleRound As New Collection
Dim ListCoordCurrent As New Collection
Dim ListCoordRound As New Collection
Dim CoordArrayCurrent(0 To 5) As Variant
Dim CoordArrayRound(0 To 5) As Variant
Dim MinDistance As Double
Dim DiffDepth As Double
Dim ListIntersection As New Collection
Dim ListIntersectionTeam As New Collection
Dim ListIntersectionTeamUniques As New Collection


num_row = Worksheets("База перспективы").Cells(Rows.Count, 2).End(xlUp).Row
start_row = 5

MinDistance = Worksheets("База перспективы").Cells(1, 3)
DiffDepth = Worksheets("База перспективы").Cells(2, 3)

Set FieldColumn = Worksheets("База перспективы").Range("B5:B" & num_row)

'Worksheets("База перспективы").Range("G5:L" & num_row).NumberFormat = "0.0"

On Error Resume Next
For Each CellField In FieldColumn
   FieldUniques.Add CellField.Value, CStr(CellField.Value)
Next CellField
On Error GoTo 0

' Добавление пропущенного заголовка
If IsEmpty(Worksheets("База перспективы").Cells(4, 19)) Then
    Worksheets("База перспективы").Cells(4, 19) = "Z"
End If

' Добавление заголовка столбца
Worksheets("База перспективы").Cells(3, 20) = "Пересечения"
Worksheets("База перспективы").Cells(3, 20).Font.Bold = True
Worksheets("База перспективы").Cells(3, 20).Font.Italic = True
Worksheets("База перспективы").Cells(4, 20) = "скв (команда/куст/объект)"
Worksheets("База перспективы").Cells(4, 20).Font.Bold = True
Worksheets("База перспективы").Cells(4, 20).Font.Italic = True
Worksheets("База перспективы").Cells(4, 20).Interior.Color = RGB(189, 215, 238)
Worksheets("База перспективы").Range("T5:T" & num_row).HorizontalAlignment = xlLeft ' Выравнивание по левому краю


' Создание листа "Статистика"
Dim NewSh As Worksheet
On Error Resume Next
Set NewSh = Sheets("Статистика")
Application.DisplayAlerts = False
If Err <> 0 Then
    Sheets.Add(After:=Sheets("База перспективы")).Name = "Статистика"
Else
    NewSh.Delete
    Sheets.Add(After:=Sheets("База перспективы")).Name = "Статистика"
End If
Application.DisplayAlerts = True
'If NewSh Is Nothing Then Sheets.Add(After:=Sheets("База перспективы")).Name = "Статистика"
Worksheets("Статистика").Cells(1, 1) = "Месторождение"
Worksheets("Статистика").Cells(1, 2) = "Количество пересечений"
Worksheets("Статистика").Cells(1, 3) = "Пересекающиеся команды"


num_field = 1

For Each Field In FieldUniques
    
    Debug.Print (Field)

    num_intersection = 0 ' Количество пересечений на месторождении
    start_hole_first = 7
    start_hole_second = 14
    
    ' Обход текущих скважин
    For i = start_row To num_row
        'Debug.Print (i)
        
        ' Проверка строки на нужное месторождение
        If Worksheets("База перспективы").Cells(i, 2) = Field Then
        
            ' Пропускаем pl
            If Worksheets("База перспективы").Cells(i, 5) = "pl" Then
                Worksheets("База перспективы").Cells(i, 20) = ""
            Else
            
                ColumnStartHoleCurrent.Add start_hole_first ' Начало первого ствола
                
                ' Проверка есть ли второй ствол
                If Not IsEmpty(Worksheets("База перспективы").Cells(i, start_hole_second)) Then
                    ColumnStartHoleCurrent.Add start_hole_second ' Начало второго ствола
                End If
                
                ' Обход стволов скважины
                For Each start_hole In ColumnStartHoleCurrent
                    x1_current = Worksheets("База перспективы").Cells(i, start_hole)
                    y1_current = Worksheets("База перспективы").Cells(i, start_hole + 1)
                    z1_current = Worksheets("База перспективы").Cells(i, start_hole + 2)
                    x3_current = Worksheets("База перспективы").Cells(i, start_hole + 3)
                    y3_current = Worksheets("База перспективы").Cells(i, start_hole + 4)
                    z3_current = Worksheets("База перспективы").Cells(i, start_hole + 5)
                    
                    'For num_coord = 0 To 5
                        'ListCoord.Add Worksheets("База перспективы").Cells(i, start_hole + num_coord)
                        
                        'If ListCoord.Item(num_coord + 1) = "-" Then
                            'ListCoord(num_coord + 1) = ""
                            'Worksheets("База перспективы").Cells(i, start_hole + num_coord) = ListCoord(num_coord + 1)
                        'ElseIf ListCoord(num_coord + 1) < 0 Then
                            'ListCoord(num_coord + 1) = ListCoord(num_coord + 1) * (-1)
                            'Worksheets("База перспективы").Cells(i, start_hole + num_coord) = ListCoord(num_coord + 1)
                        'End If
                                        
                    'Next num_coord
                    
                    
                    For num_coord = LBound(CoordArrayCurrent) To UBound(CoordArrayCurrent)
                        CoordArrayCurrent(num_coord) = Worksheets("База перспективы").Cells(i, start_hole + num_coord)
                        
                        If CoordArrayCurrent(num_coord) = "-" Or CoordArrayCurrent(num_coord) = 0 Then
                            CoordArrayCurrent(num_coord) = Empty
                            Worksheets("База перспективы").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        ElseIf CoordArrayCurrent(num_coord) < 0 Then
                            CoordArrayCurrent(num_coord) = CDbl(CoordArrayCurrent(num_coord) * (-1))
                            Worksheets("База перспективы").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        ElseIf TypeName(CoordArrayCurrent(num_coord)) = "String" Then
                            CoordArrayCurrent(num_coord) = CDbl(Worksheets("База перспективы").Cells(i, start_hole + num_coord))
                            Worksheets("База перспективы").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        End If
                        
                    Next num_coord
                    
                    ' Обход окружающих скважин
                    For j = start_row To num_row
                        'Debug.Print (j)
                    
                        ' Не проверять эту же скважину
                        If Not i = j Then
            
                            ' Проверка строки на нужное месторождение
                            If Worksheets("База перспективы").Cells(j, 2) = Field Then
                            
                                ' Пропускаем pl
                                If Worksheets("База перспективы").Cells(j, 5) = "pl" Then
                                    Worksheets("База перспективы").Cells(j, 20) = ""
                                Else
                                    ColumnStartHoleRound.Add start_hole_first ' Начало первого ствола
                        
                                    ' Проверка есть ли второй ствол
                                    If Not IsEmpty(Worksheets("База перспективы").Cells(j, start_hole_second)) Then
                                        ColumnStartHoleRound.Add start_hole_second ' Начало второго ствола
                                    End If
                                    
                                    ' Обход стволов скважины
                                    For Each start_hole_round In ColumnStartHoleRound
                                        
                                        'For num_coord = 0 To 5
                                            'ListCoord.Add Worksheets("База перспективы").Cells(j, start_hole_round + num_coord)
                                            
                                            'If ListCoord.Item(num_coord + 1) = "-" Then
                                                'ListCoord(num_coord + 1) = ""
                                                'Worksheets("База перспективы").Cells(j, start_hole_round + num_coord) = ListCoord(num_coord + 1)
                                            'ElseIf ListCoord(num_coord + 1) < 0 Then
                                                'ListCoord(num_coord + 1) = ListCoord(num_coord + 1) * (-1)
                                                'Worksheets("База перспективы").Cells(j, start_hole_round + num_coord) = ListCoord(num_coord + 1)
                                            'End If
                                                            
                                        'Next num_coord
                                        
                                        For num_coord = LBound(CoordArrayRound) To UBound(CoordArrayRound)
                                            CoordArrayRound(num_coord) = Worksheets("База перспективы").Cells(j, start_hole_round + num_coord)
                                            
                                            If CoordArrayRound(num_coord) = "-" Or CoordArrayRound(num_coord) = 0 Then
                                                CoordArrayRound(num_coord) = Empty
                                                Worksheets("База перспективы").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            ElseIf CoordArrayRound(num_coord) < 0 Then
                                                CoordArrayRound(num_coord) = CDbl(CoordArrayRound(num_coord) * (-1))
                                                Worksheets("База перспективы").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            ElseIf TypeName(CoordArrayRound(num_coord)) = "String" Then
                                                CoordArrayRound(num_coord) = CDbl(Worksheets("База перспективы").Cells(j, start_hole_round + num_coord))
                                                Worksheets("База перспективы").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            End If
                                            
                                        Next num_coord
                                        
                                        
                                        ' Проверка принадлежности к одному объекту по абс.отметка
                                        If Abs(CoordArrayCurrent(2) - CoordArrayRound(2)) < DiffDepth Or CoordArrayCurrent(2) = "" Or CoordArrayRound(2) = "" Then
                                            
                                            ' Проверка на пересечение отрезков скважин
                                            If functions.Intersect(CoordArrayCurrent(0), CoordArrayCurrent(1), CoordArrayCurrent(3), CoordArrayCurrent(4), _
                                            CoordArrayRound(0), CoordArrayRound(1), CoordArrayRound(3), CoordArrayRound(4)) Then
                                                num_intersection = num_intersection + 1
                                                ListIntersection.Add CStr(Worksheets("База перспективы").Cells(j, 6)) + " (" + CStr(Worksheets("База перспективы").Cells(j, 1)) + _
                                                                "/ " + CStr(Worksheets("База перспективы").Cells(j, 3)) + "/ " + CStr(Worksheets("База перспективы").Cells(j, 4)) + ")"
                                                
                                                ListIntersectionTeam.Add CStr(Worksheets("База перспективы").Cells(i, 1))
                                                ListIntersectionTeam.Add CStr(Worksheets("База перспективы").Cells(j, 1))
                                                
                                            ElseIf functions.MinLength(CoordArrayCurrent(0), CoordArrayCurrent(1), CoordArrayCurrent(3), CoordArrayCurrent(4), _
                                            CoordArrayRound(0), CoordArrayRound(1), CoordArrayRound(3), CoordArrayRound(4)) < MinDistance Then
                                                num_intersection = num_intersection + 1
                                                ListIntersection.Add CStr(Worksheets("База перспективы").Cells(j, 6)) + " (" + CStr(Worksheets("База перспективы").Cells(j, 1)) + _
                                                                "/ " + CStr(Worksheets("База перспективы").Cells(j, 3)) + "/ " + CStr(Worksheets("База перспективы").Cells(j, 4)) + ")"
                                                ListIntersectionTeam.Add CStr(Worksheets("База перспективы").Cells(i, 1))
                                                ListIntersectionTeam.Add CStr(Worksheets("База перспективы").Cells(j, 1))
                                                                                    
                                            End If
                                                                     
                                        End If
                                        
                                        Erase CoordArrayRound
                                    Next start_hole_round
                                End If
                            End If
                        End If
                        Set ColumnStartHoleRound = New Collection
                    Next j
                    
                    If ListIntersection.Count > 0 Then
                        Worksheets("База перспективы").Cells(i, 20) = Join(functions.CollectionToArray(ListIntersection), ",")
                    Else
                        Worksheets("База перспективы").Cells(i, 20) = ""
                    End If
                        
        
                    'Set ListCoord = New Collection
                    Erase CoordArrayCurrent
                
                Next start_hole
            End If
        End If
        Set ColumnStartHoleCurrent = New Collection
        Set ListIntersection = New Collection
    Next i
    
    
    On Error Resume Next
    For Each Team In ListIntersectionTeam
        ListIntersectionTeamUniques.Add Team, CStr(Team)
    Next Team
    On Error GoTo 0
    
    ' Заполнение листа "Статистика"
    Worksheets("Статистика").Cells(num_field + 1, 1) = FieldUniques.Item(num_field)
    Worksheets("Статистика").Cells(num_field + 1, 2) = num_intersection / 2
    If ListIntersectionTeamUniques.Count > 0 Then
        Worksheets("Статистика").Cells(num_field + 1, 3) = Join(functions.CollectionToArray(ListIntersectionTeamUniques), ", ")
    Else
        Worksheets("Статистика").Cells(num_field + 1, 3) = ""
    End If
    
    num_field = num_field + 1
    
    Set ListIntersectionTeam = New Collection
    Set ListIntersectionTeamUniques = New Collection
    
Next Field



End Sub


