Attribute VB_Name = "Volt_Analysis"
Option Explicit
Public Volt_Schedule As Worksheet
Public SubstationSheet As Worksheet
Public kV_High As Double
Public kV_Low As Double
Public Num_kV_Rows As Integer
Public kV_Row As Integer
Public Ref_Voltage As Variant
Public Q1_2016 As Long
Public Q2_2016 As Long
Public Q3_2016 As Long
Public Q4_2016 As Long
Public Q1_2017 As Long
Public Q2_2017 As Long
Public Quarter_Lengths(0 To 5) As Variant
Public Const num_data_points = 787620

Sub Main()
    Dim NumWorksheets As Integer
    Dim DataPoint As Long
    Dim VoltagesChecked As Boolean
    Dim Description As String
    Dim current_date As Date
    
    Application.ScreenUpdating = False
    
    Num_kV_Rows = 90
    
    Set Volt_Schedule = Sheets("Volt Schedules")
    
    NumWorksheets = Worksheets.Count
    ReDim Runtimes(NumWorksheets + 1)
    
    VoltagesChecked = False
    
    'Use this block to run on all sheets in workbook
    For Each SubstationSheet In Worksheets
        'Don't run script on VAR Schedules and Volt Schedules
        If SubstationSheet.Name <> "VAR Schedules" And SubstationSheet.Name <> "Volt Schedules" And SubstationSheet.Name <> "PivotTable" Then

            'Code for each Worksheet here ---------------------------------------------
            Call Initialize(SubstationSheet)

            Call RemoveOutliers(SubstationSheet)
            
            'Find the row in the schedules corresponding to the current substation
            'Assumes the name of the sheet corresponds to the substation name
            kV_Row = GetRow(Volt_Schedule, Num_kV_Rows)

            'Only check the values if we found the substation name in the schedules
            If kV_Row <> -1 Then
                'Determine if it is an exact voltage, a range, or load dependent
                Description = Volt_Schedule.Range("B1").Offset(kV_Row - 1, 0)
                For DataPoint = 1 To num_data_points
                    Call CheckVoltage(SubstationSheet, DataPoint, Description)

                    'Count the number of days in each quarter
                    current_date = SubstationSheet.Range("A1").Offset(DataPoint, 0)
                    If Application.RoundUp(Month(current_date) / 3, 0) = 1 And Year(current_date) = 2016 Then
                        Q1_2016 = Q1_2016 + 1
                    ElseIf Application.RoundUp(Month(current_date) / 3, 0) = 2 And Year(current_date) = 2016 Then
                        Q2_2016 = Q2_2016 + 1
                    ElseIf Application.RoundUp(Month(current_date) / 3, 0) = 3 And Year(current_date) = 2016 Then
                        Q3_2016 = Q3_2016 + 1
                    ElseIf Application.RoundUp(Month(current_date) / 3, 0) = 4 And Year(current_date) = 2016 Then
                        Q4_2016 = Q4_2016 + 1
                    ElseIf Application.RoundUp(Month(current_date) / 3, 0) = 1 And Year(current_date) = 2017 Then
                        Q1_2017 = Q1_2017 + 1
                    ElseIf Application.RoundUp(Month(current_date) / 3, 0) = 2 And Year(current_date) = 2017 Then
                        Q2_2017 = Q2_2017 + 1
                    End If
                Next

                Quarter_Lengths(0) = Q1_2016
                Quarter_Lengths(1) = Q2_2016
                Quarter_Lengths(2) = Q3_2016
                Quarter_Lengths(3) = Q4_2016
                Quarter_Lengths(4) = Q1_2017
                Quarter_Lengths(5) = Q2_2017

                VoltagesChecked = True
            End If

            'Only compute the annual summary if we actually checked the voltages
            If VoltagesChecked Then
                Call VoltageSummary(SubstationSheet)
            End If
            '--------------------------------------------------------------------------
        End If
    Next
    Application.ScreenUpdating = True
End Sub

Sub Initialize(ws As Worksheet)
    Dim i As Long
    Dim current_date As Date
    
    Q1_2016 = 0
    Q2_2016 = 0
    Q3_2016 = 0
    Q4_2016 = 0
    Q1_2017 = 0
    Q2_2017 = 0
    
    With ws
        .Range("A1") = "Date"
        .Range("B1") = "Time"
        .Range("C1") = "MW"
        .Range("D1") = "MVAR"
        .Range("E1") = "In Range"
        .Range("F1") = "> -10 MVAR"
        .Range("G1") = "> -20 MVAR"
        .Range("H1") = "<= -20 MVAR"
        .Range("I1") = "Upper VAR Threshold"
        .Range("J1") = "Lower VAR Threshold"
        .Range("K1") = "Actual Difference"
        .Range("L1") = "kV"
        .Range("M1") = "Reference"
        .Range("N1") = "Difference"
        .Range("O1") = "In Range"
        .Range("P1") = "Above Range"
        .Range("Q1") = "Below Range"
        
        'Set the cells for the annual summary
        .Range("A787625").Offset(1, 0) = "Q1 2016"
        .Range("A787625").Offset(2, 0) = "Q2 2016"
        .Range("A787625").Offset(3, 0) = "Q3 2016"
        .Range("A787625").Offset(4, 0) = "Q4 2016"
        .Range("A787625").Offset(5, 0) = "Q1 2017"
        .Range("A787625").Offset(6, 0) = "Q2 2017"
        
        .Range("A787625").Offset(0, 1) = "# in Quarter"
        .Range("A787625").Offset(0, 1).Font.Bold = True
        .Range("A787625").Offset(0, 2) = "# In Range"
        .Range("A787625").Offset(0, 2).Font.Bold = True
        .Range("A787625").Offset(0, 3) = "% In Range"
        .Range("A787625").Offset(0, 3).Font.Bold = True
        .Range("A787625").Offset(0, 4) = "# Above Range"
        .Range("A787625").Offset(0, 4).Font.Bold = True
        .Range("A787625").Offset(0, 5) = "% Above Range"
        .Range("A787625").Offset(0, 5).Font.Bold = True
        .Range("A787625").Offset(0, 6) = "# Below Range"
        .Range("A787625").Offset(0, 6).Font.Bold = True
        .Range("A787625").Offset(0, 7) = "% Below Range"
        .Range("A787625").Offset(0, 7).Font.Bold = True
    End With
End Sub

Sub RemoveOutliers(ws As Worksheet)
    Dim i As Long
    With ws
        For i = 1 To num_data_points
            If .Range("L1").Offset(i, 0) < 500 Then
                .Range("L1").Offset(i, 0) = .Range("L1").Offset(i - 1, 0)
            End If
        Next
    End With
End Sub

Function GetRow(ReferenceSheet As Worksheet, NumRows As Integer) As Integer
    'This function returns the row number in column A of ReferenceSheet
    '    where the Sheet name occurs. If the name is not present, return -1.
    
    Dim Substation_Name As String
    Dim row As Integer
    Dim found As Boolean
    
    Substation_Name = SubstationSheet.Name
    found = False
    For row = 1 To NumRows
        If Substation_Name = ReferenceSheet.Range("A1").Offset(row, 0) And Not found Then
            GetRow = row + 1
            found = True
        ElseIf IsEmpty(ReferenceSheet.Range("A1").Offset(row, 0)) Then
            Exit For
        End If
    Next
    
    If Not found Then
        GetRow = -1
    End If
End Function

Sub CheckVoltage(ws As Worksheet, DataPoint As Long, Description As String)
    'This procedure checks if the substation voltages are:
    '       within 1% of the reference voltage
    '       within 1.5% but outside 1% of the reference voltage
    '       outside 1.5% of the reference voltage
    'and writes a 1 in the cell if its not in the range and 0 if it is in the range.
    '(1 for violation, 0 for no violation)
    Dim Voltage As Double
    Dim Difference As Double
    
    Dim High_V As Double
    Dim Low_V As Double
    
    With ws
        Voltage = .Range("L1").Offset(DataPoint, 0)
        
        If Description = "range" Then
            'Write the reference voltage range in cell
            Low_V = Volt_Schedule.Range("E1").Offset(kV_Row - 1, 0)
            High_V = Volt_Schedule.Range("F1").Offset(kV_Row - 1, 0)
            .Range("M1").Offset(DataPoint, 0) = "[" & Low_V & ", " & High_V & "]"
            
            'Difference is only non-zero if out of range
            If Voltage < Low_V Then
                Difference = Voltage - Low_V
            ElseIf Voltage > High_V Then
                Difference = Voltage - High_V
            Else
                Difference = 0
            End If
            
            'Write a 1 in the cell there was a violation; that is, the voltage is outside the allowable range
            If Voltage > High_V Then
                .Range("O1").Offset(DataPoint, 0) = 0
                .Range("P1").Offset(DataPoint, 0) = 1
                .Range("Q1").Offset(DataPoint, 0) = 0
            ElseIf Voltage < Low_V Then
                .Range("O1").Offset(DataPoint, 0) = 0
                .Range("P1").Offset(DataPoint, 0) = 0
                .Range("Q1").Offset(DataPoint, 0) = 1
            ElseIf Voltage <= High_V And Voltage >= Low_V Then
                .Range("O1").Offset(DataPoint, 0) = 1
                .Range("P1").Offset(DataPoint, 0) = 0
                .Range("Q1").Offset(DataPoint, 0) = 0
            End If
        End If
        .Range("N1").Offset(DataPoint, 0) = Difference
    End With
End Sub

Sub VoltageSummary(ws As Worksheet)
    Dim In_Range As Long
    Dim Above_Range As Long
    Dim Below_Range As Long
    Dim num_data As Long
    Dim index As Integer
    Dim QLength As Long
    
    num_data = 0
    
    With ws
        For index = 0 To 5
            QLength = Quarter_Lengths(index)
            .Range("A787625").Offset(index + 1, 1) = QLength
            If QLength <> 0 Then
                In_Range = Application.Sum(.Range("O2").Offset(num_data, 0).Resize(QLength))
                Above_Range = Application.Sum(.Range("P2").Offset(num_data, 0).Resize(QLength))
                Below_Range = Application.Sum(.Range("Q2").Offset(num_data, 0).Resize(QLength))
                .Range("A787625").Offset(index + 1, 2) = In_Range
                .Range("A787625").Offset(index + 1, 3) = Format(In_Range / QLength, "Percent")
                .Range("A787625").Offset(index + 1, 4) = Above_Range
                .Range("A787625").Offset(index + 1, 5) = Format(Above_Range / QLength, "Percent")
                .Range("A787625").Offset(index + 1, 6) = Below_Range
                .Range("A787625").Offset(index + 1, 7) = Format(Below_Range / QLength, "Percent")
                num_data = num_data + QLength
            End If
        Next
    End With
End Sub

'For creating a PivotTable
Sub CreateCategories()
    Dim i As Long
    Dim current_date As Date
    Dim current_time As String
    
    For Each SubstationSheet In Worksheets
        'Don't run script on VAR Schedules and Volt Schedules
        If SubstationSheet.Name <> "VAR Schedules" And SubstationSheet.Name <> "Volt Schedules" Then
            For i = 1 To 787620
                current_date = SubstationSheet.Range("A1").Offset(i, 0)
                current_time = SubstationSheet.Range("B1").Offset(i, 0)
                SubstationSheet.Range("R1").Offset(i, 0) = Year(current_date)
                SubstationSheet.Range("S1").Offset(i, 0) = Month(current_date)
                SubstationSheet.Range("T1").Offset(i, 0) = Weekday(current_date)
                SubstationSheet.Range("U1").Offset(i, 0) = Hour(current_time)
            Next
        End If
    Next
End Sub


