Attribute VB_Name = "main"
Sub SearchIntersection()
'���������� ����� �� �����

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


num_row = Worksheets("���� �����������").Cells(Rows.Count, 2).End(xlUp).Row
start_row = 5

MinDistance = Worksheets("���� �����������").Cells(1, 3)
DiffDepth = Worksheets("���� �����������").Cells(2, 3)

Set FieldColumn = Worksheets("���� �����������").Range("B5:B" & num_row)

'Worksheets("���� �����������").Range("G5:L" & num_row).NumberFormat = "0.0"

On Error Resume Next
For Each CellField In FieldColumn
   FieldUniques.Add CellField.Value, CStr(CellField.Value)
Next CellField
On Error GoTo 0

' ���������� ������������ ���������
If IsEmpty(Worksheets("���� �����������").Cells(4, 19)) Then
    Worksheets("���� �����������").Cells(4, 19) = "Z"
End If

' ���������� ��������� �������
Worksheets("���� �����������").Cells(3, 20) = "�����������"
Worksheets("���� �����������").Cells(3, 20).Font.Bold = True
Worksheets("���� �����������").Cells(3, 20).Font.Italic = True
Worksheets("���� �����������").Cells(4, 20) = "��� (�������/����/������)"
Worksheets("���� �����������").Cells(4, 20).Font.Bold = True
Worksheets("���� �����������").Cells(4, 20).Font.Italic = True
Worksheets("���� �����������").Cells(4, 20).Interior.Color = RGB(189, 215, 238)
Worksheets("���� �����������").Range("T5:T" & num_row).HorizontalAlignment = xlLeft ' ������������ �� ������ ����


' �������� ����� "����������"
Dim NewSh As Worksheet
On Error Resume Next
Set NewSh = Sheets("����������")
Application.DisplayAlerts = False
If Err <> 0 Then
    Sheets.Add(After:=Sheets("���� �����������")).Name = "����������"
Else
    NewSh.Delete
    Sheets.Add(After:=Sheets("���� �����������")).Name = "����������"
End If
Application.DisplayAlerts = True
'If NewSh Is Nothing Then Sheets.Add(After:=Sheets("���� �����������")).Name = "����������"
Worksheets("����������").Cells(1, 1) = "�������������"
Worksheets("����������").Cells(1, 2) = "���������� �����������"
Worksheets("����������").Cells(1, 3) = "�������������� �������"


num_field = 1

For Each Field In FieldUniques
    
    Debug.Print (Field)

    num_intersection = 0 ' ���������� ����������� �� �������������
    start_hole_first = 7
    start_hole_second = 14
    
    ' ����� ������� �������
    For i = start_row To num_row
        'Debug.Print (i)
        
        ' �������� ������ �� ������ �������������
        If Worksheets("���� �����������").Cells(i, 2) = Field Then
        
            ' ���������� pl
            If Worksheets("���� �����������").Cells(i, 5) = "pl" Then
                Worksheets("���� �����������").Cells(i, 20) = ""
            Else
            
                ColumnStartHoleCurrent.Add start_hole_first ' ������ ������� ������
                
                ' �������� ���� �� ������ �����
                If Not IsEmpty(Worksheets("���� �����������").Cells(i, start_hole_second)) Then
                    ColumnStartHoleCurrent.Add start_hole_second ' ������ ������� ������
                End If
                
                ' ����� ������� ��������
                For Each start_hole In ColumnStartHoleCurrent
                    x1_current = Worksheets("���� �����������").Cells(i, start_hole)
                    y1_current = Worksheets("���� �����������").Cells(i, start_hole + 1)
                    z1_current = Worksheets("���� �����������").Cells(i, start_hole + 2)
                    x3_current = Worksheets("���� �����������").Cells(i, start_hole + 3)
                    y3_current = Worksheets("���� �����������").Cells(i, start_hole + 4)
                    z3_current = Worksheets("���� �����������").Cells(i, start_hole + 5)
                    
                    'For num_coord = 0 To 5
                        'ListCoord.Add Worksheets("���� �����������").Cells(i, start_hole + num_coord)
                        
                        'If ListCoord.Item(num_coord + 1) = "-" Then
                            'ListCoord(num_coord + 1) = ""
                            'Worksheets("���� �����������").Cells(i, start_hole + num_coord) = ListCoord(num_coord + 1)
                        'ElseIf ListCoord(num_coord + 1) < 0 Then
                            'ListCoord(num_coord + 1) = ListCoord(num_coord + 1) * (-1)
                            'Worksheets("���� �����������").Cells(i, start_hole + num_coord) = ListCoord(num_coord + 1)
                        'End If
                                        
                    'Next num_coord
                    
                    
                    For num_coord = LBound(CoordArrayCurrent) To UBound(CoordArrayCurrent)
                        CoordArrayCurrent(num_coord) = Worksheets("���� �����������").Cells(i, start_hole + num_coord)
                        
                        If CoordArrayCurrent(num_coord) = "-" Or CoordArrayCurrent(num_coord) = 0 Then
                            CoordArrayCurrent(num_coord) = Empty
                            Worksheets("���� �����������").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        ElseIf CoordArrayCurrent(num_coord) < 0 Then
                            CoordArrayCurrent(num_coord) = CDbl(CoordArrayCurrent(num_coord) * (-1))
                            Worksheets("���� �����������").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        ElseIf TypeName(CoordArrayCurrent(num_coord)) = "String" Then
                            CoordArrayCurrent(num_coord) = CDbl(Worksheets("���� �����������").Cells(i, start_hole + num_coord))
                            Worksheets("���� �����������").Cells(i, start_hole + num_coord) = CoordArrayCurrent(num_coord)
                        End If
                        
                    Next num_coord
                    
                    ' ����� ���������� �������
                    For j = start_row To num_row
                        'Debug.Print (j)
                    
                        ' �� ��������� ��� �� ��������
                        If Not i = j Then
            
                            ' �������� ������ �� ������ �������������
                            If Worksheets("���� �����������").Cells(j, 2) = Field Then
                            
                                ' ���������� pl
                                If Worksheets("���� �����������").Cells(j, 5) = "pl" Then
                                    Worksheets("���� �����������").Cells(j, 20) = ""
                                Else
                                    ColumnStartHoleRound.Add start_hole_first ' ������ ������� ������
                        
                                    ' �������� ���� �� ������ �����
                                    If Not IsEmpty(Worksheets("���� �����������").Cells(j, start_hole_second)) Then
                                        ColumnStartHoleRound.Add start_hole_second ' ������ ������� ������
                                    End If
                                    
                                    ' ����� ������� ��������
                                    For Each start_hole_round In ColumnStartHoleRound
                                        
                                        'For num_coord = 0 To 5
                                            'ListCoord.Add Worksheets("���� �����������").Cells(j, start_hole_round + num_coord)
                                            
                                            'If ListCoord.Item(num_coord + 1) = "-" Then
                                                'ListCoord(num_coord + 1) = ""
                                                'Worksheets("���� �����������").Cells(j, start_hole_round + num_coord) = ListCoord(num_coord + 1)
                                            'ElseIf ListCoord(num_coord + 1) < 0 Then
                                                'ListCoord(num_coord + 1) = ListCoord(num_coord + 1) * (-1)
                                                'Worksheets("���� �����������").Cells(j, start_hole_round + num_coord) = ListCoord(num_coord + 1)
                                            'End If
                                                            
                                        'Next num_coord
                                        
                                        For num_coord = LBound(CoordArrayRound) To UBound(CoordArrayRound)
                                            CoordArrayRound(num_coord) = Worksheets("���� �����������").Cells(j, start_hole_round + num_coord)
                                            
                                            If CoordArrayRound(num_coord) = "-" Or CoordArrayRound(num_coord) = 0 Then
                                                CoordArrayRound(num_coord) = Empty
                                                Worksheets("���� �����������").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            ElseIf CoordArrayRound(num_coord) < 0 Then
                                                CoordArrayRound(num_coord) = CDbl(CoordArrayRound(num_coord) * (-1))
                                                Worksheets("���� �����������").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            ElseIf TypeName(CoordArrayRound(num_coord)) = "String" Then
                                                CoordArrayRound(num_coord) = CDbl(Worksheets("���� �����������").Cells(j, start_hole_round + num_coord))
                                                Worksheets("���� �����������").Cells(j, start_hole_round + num_coord) = CoordArrayRound(num_coord)
                                            End If
                                            
                                        Next num_coord
                                        
                                        
                                        ' �������� �������������� � ������ ������� �� ���.�������
                                        If Abs(CoordArrayCurrent(2) - CoordArrayRound(2)) < DiffDepth Or CoordArrayCurrent(2) = "" Or CoordArrayRound(2) = "" Then
                                            
                                            ' �������� �� ����������� �������� �������
                                            If functions.Intersect(CoordArrayCurrent(0), CoordArrayCurrent(1), CoordArrayCurrent(3), CoordArrayCurrent(4), _
                                            CoordArrayRound(0), CoordArrayRound(1), CoordArrayRound(3), CoordArrayRound(4)) Then
                                                num_intersection = num_intersection + 1
                                                ListIntersection.Add CStr(Worksheets("���� �����������").Cells(j, 6)) + " (" + CStr(Worksheets("���� �����������").Cells(j, 1)) + _
                                                                "/ " + CStr(Worksheets("���� �����������").Cells(j, 3)) + "/ " + CStr(Worksheets("���� �����������").Cells(j, 4)) + ")"
                                                
                                                ListIntersectionTeam.Add CStr(Worksheets("���� �����������").Cells(i, 1))
                                                ListIntersectionTeam.Add CStr(Worksheets("���� �����������").Cells(j, 1))
                                                
                                            ElseIf functions.MinLength(CoordArrayCurrent(0), CoordArrayCurrent(1), CoordArrayCurrent(3), CoordArrayCurrent(4), _
                                            CoordArrayRound(0), CoordArrayRound(1), CoordArrayRound(3), CoordArrayRound(4)) < MinDistance Then
                                                num_intersection = num_intersection + 1
                                                ListIntersection.Add CStr(Worksheets("���� �����������").Cells(j, 6)) + " (" + CStr(Worksheets("���� �����������").Cells(j, 1)) + _
                                                                "/ " + CStr(Worksheets("���� �����������").Cells(j, 3)) + "/ " + CStr(Worksheets("���� �����������").Cells(j, 4)) + ")"
                                                ListIntersectionTeam.Add CStr(Worksheets("���� �����������").Cells(i, 1))
                                                ListIntersectionTeam.Add CStr(Worksheets("���� �����������").Cells(j, 1))
                                                                                    
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
                        Worksheets("���� �����������").Cells(i, 20) = Join(functions.CollectionToArray(ListIntersection), ",")
                    Else
                        Worksheets("���� �����������").Cells(i, 20) = ""
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
    
    ' ���������� ����� "����������"
    Worksheets("����������").Cells(num_field + 1, 1) = FieldUniques.Item(num_field)
    Worksheets("����������").Cells(num_field + 1, 2) = num_intersection / 2
    If ListIntersectionTeamUniques.Count > 0 Then
        Worksheets("����������").Cells(num_field + 1, 3) = Join(functions.CollectionToArray(ListIntersectionTeamUniques), ", ")
    Else
        Worksheets("����������").Cells(num_field + 1, 3) = ""
    End If
    
    num_field = num_field + 1
    
    Set ListIntersectionTeam = New Collection
    Set ListIntersectionTeamUniques = New Collection
    
Next Field



End Sub


