Attribute VB_Name = "VoltVAR_Analysis"
Option Explicit
Public MWOK As Integer
Public MVAROK As Integer
Public MVAROK1 As Integer
Public MVAROK2 As Integer
Public MVAROK3 As Integer
Public SnapDown As Long
Public MWRight As Integer
Public NameY As Integer
Public NameX As Integer
Public SnapData As String
Public HBAD As Variant
Public LBAD As Variant
Public MVARValPub As Variant
Public HighM As Integer
Public LowM As Integer
Public ws As Worksheet
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
    Call VARAnalysis
    Call VoltAnalysis
End Sub

Sub VARAnalysis()

    Dim QLength As Long
    Dim QLength2 As Long
    Dim QLength3 As Long
    Dim QLength4 As Long
    Dim QLength5 As Long
    Dim QLength6 As Long
    
    SnapDown = 0
    NameY = 0
    QLength = 130980                '131040 Daylight savings 'These need to be double checked
    QLength2 = 131040
    QLength3 = 132480
    QLength4 = 132540
    QLength5 = 129540       'Daylight savings
    QLength6 = 131040
    
    SnapData = Application.ActiveSheet.Name         'ENTER NAME HERE
    Set ws = Worksheets(SnapData)
    
    Call FindNameY
    
    For SnapDown = 1 To num_data_points
    
        Call MWCheck
        Call MVARCheck
        
        If MWOK And MVAROK = 1 Then
            ws.Range("E1").Offset(SnapDown, 0) = 1
        Else
            ws.Range("E1").Offset(SnapDown, 0) = 0
        End If
        
        If MWOK And MVAROK1 = 1 Then
            ws.Range("F1").Offset(SnapDown, 0) = 1
        Else
            ws.Range("F1").Offset(SnapDown, 0) = 0
        End If
        
        If MWOK And MVAROK2 = 1 Then
            ws.Range("G1").Offset(SnapDown, 0) = 1
        Else
            ws.Range("G1").Offset(SnapDown, 0) = 0
        End If
        
        If MWOK And MVAROK3 = 1 Then
            ws.Range("H1").Offset(SnapDown, 0) = 1
        Else
            ws.Range("H1").Offset(SnapDown, 0) = 0
        End If
        
        Call FindActualDiff
    
    Next SnapDown
    
    ws.Range("T1") = "In Range"
    ws.Range("T1").Font.Bold = True
    ws.Range("U1") = "> -10 MVAR"
    ws.Range("U1").Font.Bold = True
    ws.Range("V1") = "> -20 MVAR"
    ws.Range("V1").Font.Bold = True
    ws.Range("W1") = "<= -20 MVAR"
    ws.Range("W1").Font.Bold = True
    
    'Calculate the sum of 1's Q1
    ws.Range("T2").Offset(0, -1) = "Q1 2016 Total"
    ws.Range("T2").Offset(0, 0) = Application.Sum(Range("E1:E130979").Offset(0, 0))
    ws.Range("T2").Offset(0, 1) = Application.Sum(Range("E1:E130979").Offset(0, 1))
    ws.Range("T2").Offset(0, 2) = Application.Sum(Range("E1:E130979").Offset(0, 2))
    ws.Range("T2").Offset(0, 3) = Application.Sum(Range("E1:E130979").Offset(0, 3))
    
    'Find the percent of 1's to 0's Q1
    ws.Range("T2").Offset(1, -1) = "Q1 2016 Percentage"
    ws.Range("T2").Offset(1, 0) = ws.Range("T2").Offset(0, 0) / QLength
    ws.Range("T2").Offset(1, 1) = ws.Range("T2").Offset(0, 1) / QLength
    ws.Range("T2").Offset(1, 2) = ws.Range("T2").Offset(0, 2) / QLength
    ws.Range("T2").Offset(1, 3) = ws.Range("T2").Offset(0, 3) / QLength
    
    'Calculate the sum of 1's Q2
    ws.Range("T2").Offset(2, -1) = "Q2 2016 Total"
    ws.Range("T2").Offset(2, 0) = Application.Sum(Range("E130980:E262019").Offset(0, 0))
    ws.Range("T2").Offset(2, 1) = Application.Sum(Range("E130980:E262019").Offset(0, 1))
    ws.Range("T2").Offset(2, 2) = Application.Sum(Range("E130980:E262019").Offset(0, 2))
    ws.Range("T2").Offset(2, 3) = Application.Sum(Range("E130980:E262019").Offset(0, 3))
    
    'Find the percent of 1's to 0's Q2
    ws.Range("T2").Offset(3, -1) = "Q2 2016 Percentage"
    ws.Range("T2").Offset(3, 0) = ws.Range("T4").Offset(0, 0) / QLength2
    ws.Range("T2").Offset(3, 1) = ws.Range("T4").Offset(0, 1) / QLength2
    ws.Range("T2").Offset(3, 2) = ws.Range("T4").Offset(0, 2) / QLength2
    ws.Range("T2").Offset(3, 3) = ws.Range("T4").Offset(0, 3) / QLength2
    
    'Calculate the sum of 1's Q3
    ws.Range("T2").Offset(4, -1) = "Q3 2016 Total"
    ws.Range("T2").Offset(4, 0) = Application.Sum(Range("E262020:E394499").Offset(0, 0))
    ws.Range("T2").Offset(4, 1) = Application.Sum(Range("E262020:E394499").Offset(0, 1))
    ws.Range("T2").Offset(4, 2) = Application.Sum(Range("E262020:E394499").Offset(0, 2))
    ws.Range("T2").Offset(4, 3) = Application.Sum(Range("E262020:E394499").Offset(0, 3))
     
    'Find the percent of 1's to 0's Q3
    ws.Range("T2").Offset(5, -1) = "Q3 2016 Percentage"
    ws.Range("T2").Offset(5, 0) = ws.Range("T6").Offset(0, 0) / QLength3
    ws.Range("T2").Offset(5, 1) = ws.Range("T6").Offset(0, 1) / QLength3
    ws.Range("T2").Offset(5, 2) = ws.Range("T6").Offset(0, 2) / QLength3
    ws.Range("T2").Offset(5, 3) = ws.Range("T6").Offset(0, 3) / QLength3
    
    'Calculate the sum of 1's Q4
    ws.Range("T2").Offset(6, -1) = "Q4 2016 Total"
    ws.Range("T2").Offset(6, 0) = Application.Sum(Range("E394500:E527039").Offset(0, 0))
    ws.Range("T2").Offset(6, 1) = Application.Sum(Range("E394500:E527039").Offset(0, 1))
    ws.Range("T2").Offset(6, 2) = Application.Sum(Range("E394500:E527039").Offset(0, 2))
    ws.Range("T2").Offset(6, 3) = Application.Sum(Range("E394500:E527039").Offset(0, 3))
    
    'Find the percent of 1's to 0's Q4
    ws.Range("T2").Offset(7, -1) = "Q4 2016 Percentage"
    ws.Range("T2").Offset(7, 0) = ws.Range("T8").Offset(0, 0) / QLength4
    ws.Range("T2").Offset(7, 1) = ws.Range("T8").Offset(0, 1) / QLength4
    ws.Range("T2").Offset(7, 2) = ws.Range("T8").Offset(0, 2) / QLength4
    ws.Range("T2").Offset(7, 3) = ws.Range("T8").Offset(0, 3) / QLength4
    
    'Calculate the sum of 1's Q1 again
    ws.Range("T2").Offset(8, -1) = "Q1 2017 Total"
    ws.Range("T2").Offset(8, 0) = Application.Sum(Range("E527040:E656579").Offset(0, 0))
    ws.Range("T2").Offset(8, 1) = Application.Sum(Range("E527040:E656579").Offset(0, 1))
    ws.Range("T2").Offset(8, 2) = Application.Sum(Range("E527040:E656579").Offset(0, 2))
    ws.Range("T2").Offset(8, 3) = Application.Sum(Range("E527040:E656579").Offset(0, 3))
    
    'Find the percent of 1's to 0's Q1 again
    ws.Range("T2").Offset(9, -1) = "Q1 2017 Percentage"
    ws.Range("T2").Offset(9, 0) = ws.Range("T10").Offset(0, 0) / QLength5
    ws.Range("T2").Offset(9, 1) = ws.Range("T10").Offset(0, 1) / QLength5
    ws.Range("T2").Offset(9, 2) = ws.Range("T10").Offset(0, 2) / QLength5
    ws.Range("T2").Offset(9, 3) = ws.Range("T10").Offset(0, 3) / QLength5

    'Calculate the sum of 1's Q2 again
    ws.Range("T2").Offset(10, -1) = "Q2 2017 Total"
    ws.Range("T2").Offset(10, 0) = Application.Sum(Range("E656580:E787619").Offset(0, 0))
    ws.Range("T2").Offset(10, 1) = Application.Sum(Range("E656580:E787619").Offset(0, 1))
    ws.Range("T2").Offset(10, 2) = Application.Sum(Range("E656580:E787619").Offset(0, 2))
    ws.Range("T2").Offset(10, 3) = Application.Sum(Range("E656580:E787619").Offset(0, 3))
    
    'Find the percent of 1's to 0's Q2 again
    ws.Range("T2").Offset(11, -1) = "Q2 2017 Percentage"
    ws.Range("T2").Offset(11, 0) = ws.Range("T12").Offset(0, 0) / QLength6
    ws.Range("T2").Offset(11, 1) = ws.Range("T12").Offset(0, 1) / QLength6
    ws.Range("T2").Offset(11, 2) = ws.Range("T12").Offset(0, 2) / QLength6
    ws.Range("T2").Offset(11, 3) = ws.Range("T12").Offset(0, 3) / QLength6
End Sub

Sub MWCheck()

    Dim MWVal As Variant
    Dim MWH As Variant
    Dim MWL As Variant
    Dim MWValE As Integer
    Dim MWHE As Integer
    Dim MWLE As Integer
    
    MWHE = 1
    MWLE = 1
    
    For MWRight = 0 To 10
    
        'To check to see if there is still something in the cell
        
        MWHE = IsEmpty(Worksheets("VAR Schedules").Range("D2").Offset(NameY, MWRight * 4))
        MWLE = IsEmpty(Worksheets("VAR Schedules").Range("C2").Offset(NameY, MWRight * 4))
        
        If MWHE = -1 Then
            MWH = 1100     'No higher value
        End If
          
        If MWLE = -1 Then
            MWOK = 0
            Exit Sub       'Ran out of test cases, exit sub
        End If
    
        MWVal = ws.Range("C1").Offset(SnapDown, NameX)
            'Seperate pages, don't comapre offset
        
        If MWH < 1100 Then
            MWH = Worksheets("VAR Schedules").Range("D2").Offset(NameY, MWRight * 4)
        End If
        
        MWL = Worksheets("VAR Schedules").Range("C2").Offset(NameY, MWRight * 4)
        
        If MWLE = 0 Then
            If MWVal <= MWH Then
                MWOK = 1
                Exit Sub
            Else
                MWOK = 0
            End If
        Else
            If MWVal >= MWL And MWVal <= MWH Then
                MWOK = 1
                Exit Sub
            Else
                MWOK = 0
            End If
        End If
    Next MWRight

End Sub
Sub MVARCheck()

    Dim MVARVal As Variant
    Dim MVARH As Variant
    Dim MVARL As Variant
    Dim MVARRight As Integer
    Dim MVARValE As Integer
    Dim MVARHE As Integer
    Dim MVARLE As Integer
    Dim MVARL10 As Integer
    Dim MVARL20 As Integer

        MVARVal = ws.Range("D1").Offset(SnapDown, NameX)
        MVARValPub = MVARVal
        'Seperate pages, don't comapre offset
        MVARH = Worksheets("VAR Schedules").Range("E2").Offset(NameY, MWRight * 4)
        MVARL = Worksheets("VAR Schedules").Range("F2").Offset(NameY, MWRight * 4)
        MVARL10 = MVARL - 10
        MVARL20 = MVARL - 20
        
        ws.Range("I1").Offset(SnapDown, 0) = MVARH
        ws.Range("J1").Offset(SnapDown, 0) = MVARL
        
        If MVARVal >= MVARL Then 'And MVARVal < MVARH Then
            MVAROK = 1
        Else
            MVAROK = 0
        End If

        If MVARVal >= MVARL10 And MVARVal < MVARL Then
            MVAROK1 = 1
        Else
            MVAROK1 = 0
        End If
        
        If MVARVal >= MVARL20 And MVARVal < MVARL10 Then
            MVAROK2 = 1
        Else
            MVAROK2 = 0
        End If
        
        If MVARVal < MVARL20 Then
            MVAROK3 = 1
        Else
            MVAROK3 = 0
        End If
End Sub

Sub FindNameY()

Dim CycleNames As Integer
Dim CurrentName As String

For CycleNames = 0 To 52
    CurrentName = Worksheets("VAR Schedules").Range("A2").Offset(CycleNames, 0)
    If CurrentName = SnapData Then
        NameY = CycleNames
        Exit Sub
    End If
Next CycleNames
End Sub

Sub FindActualDiff()

    Dim Ans1 As Variant
    Dim Ans2 As Variant
    Dim Skip As Integer
    
    HBAD = 0
    LBAD = 0
    HighM = 0
    LowM = 0
    Skip = 0
    
    HBAD = ws.Range("I1").Offset(SnapDown, 0)
    LBAD = ws.Range("J1").Offset(SnapDown, 0)
    
    Ans1 = HBAD - MVARValPub
    Ans2 = MVARValPub - LBAD
    
    If MVAROK = 1 Then
        ws.Range("K1").Offset(SnapDown, 0) = 0
        Skip = 1
    End If
    
    If Skip = 0 Then
        If Ans1 > Ans2 Then
            ws.Range("K1").Offset(SnapDown, 0) = Ans2
            HighM = 0
            LowM = 1
        Else
            ws.Range("K1").Offset(SnapDown, 0) = Ans1
            HighM = 1
            LowM = 0
        End If
    End If
End Sub

Sub VoltAnalysis()
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
        If SubstationSheet.Name <> "VAR Schedules" And SubstationSheet.Name <> "Volt Schedules" Then

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
                    current_date = SubstationSheet.Range("A3").Offset(DataPoint, 0)
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
        .Range("S15").Offset(1, 0) = "Q1 2016"
        .Range("S15").Offset(2, 0) = "Q2 2016"
        .Range("S15").Offset(3, 0) = "Q3 2016"
        .Range("S15").Offset(4, 0) = "Q4 2016"
        .Range("S15").Offset(5, 0) = "Q1 2017"
        .Range("S15").Offset(6, 0) = "Q2 2017"
        
        .Range("S15").Offset(0, 1) = "# in Quarter"
        .Range("S15").Offset(0, 1).Font.Bold = True
        .Range("S15").Offset(0, 2) = "# In Range"
        .Range("S15").Offset(0, 2).Font.Bold = True
        .Range("S15").Offset(0, 3) = "% In Range"
        .Range("S15").Offset(0, 3).Font.Bold = True
        .Range("S15").Offset(0, 4) = "# Above Range"
        .Range("S15").Offset(0, 4).Font.Bold = True
        .Range("S15").Offset(0, 5) = "% Above Range"
        .Range("S15").Offset(0, 5).Font.Bold = True
        .Range("S15").Offset(0, 6) = "# Below Range"
        .Range("S15").Offset(0, 6).Font.Bold = True
        .Range("S15").Offset(0, 7) = "% Below Range"
        .Range("S15").Offset(0, 7).Font.Bold = True
    End With
End Sub

Sub RemoveOutliers(ws As Worksheet)
    Dim i As Long
    Dim iter As Long
    Dim average As Double
    Dim before As Long
    Dim after As Long
    Dim outlier As Boolean
    
    outlier = False
    
    With ws
        For i = 1 To 787620
            If .Range("L2").Offset(i, 0) < 60 And Not outlier Then
                outlier = True
                before = i - 1
            End If
            
            If .Range("L2").Offset(i, 0) > 60 And outlier Then
                outlier = False
                after = i
                
                average = (.Range("L2").Offset(before, 0) + .Range("L2").Offset(after, 0)) / 2
                For iter = before + 1 To after - 1
                    .Range("L2").Offset(iter) = average
                Next iter
            End If
        Next i
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
    For row = 0 To NumRows
        If Substation_Name = ReferenceSheet.Range("A1").Offset(row, 0) And Not found Then
            GetRow = row + 1
            found = True
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
    
    Dim Delta_1 As Double
    Dim Delta_1_5 As Double
    Dim Delta_2 As Double
    
    Dim High_V As Double
    Dim Low_V As Double

    Dim Delta_2_Low As Double
    Dim Delta_2_High As Double
    
    Dim Load_Index As Integer
    Dim Load As Double
    Dim High_Load As Double
    Dim Low_Load As Double
    Dim Load_Voltage As Double
    
    With ws
        Voltage = .Range("L1").Offset(DataPoint, 0)
        
        If Description = "all times" Then
            'Write the reference voltage in cell
            Ref_Voltage = Volt_Schedule.Range("D1").Offset(kV_Row - 1, 0)
            Low_V = 0.98 * Ref_Voltage
            High_V = 1.02 * Ref_Voltage
            .Range("M1").Offset(DataPoint, 0) = "[" & Low_V & ", " & High_V & "]"
            
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
            
        ElseIf Description = "range" Then
            'Write the reference voltage range in cell
            Low_V = Volt_Schedule.Range("E1").Offset(kV_Row - 1, 0)
            High_V = Volt_Schedule.Range("F1").Offset(kV_Row - 1, 0)
            .Range("M1").Offset(DataPoint, 0) = "[" & Low_V & ", " & High_V & "]"
            
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
            
        ElseIf Description = "load dependent" Then
            High_Load = 0
            Low_Load = 0
            Load = .Range("C1").Offset(DataPoint, 0)
            
            For Load_Index = 0 To 7
                If Not IsEmpty(Volt_Schedule.Range("G1").Offset(kV_Row - 1, 3 * Load_Index)) Then
                    Low_Load = Volt_Schedule.Range("G1").Offset(kV_Row - 1, 3 * Load_Index)
                    Load_Voltage = Volt_Schedule.Range("G1").Offset(kV_Row - 1, 3 * Load_Index + 2)
                        
                    If Not IsEmpty(Volt_Schedule.Range("G1").Offset(kV_Row - 1, 3 * Load_Index) + 1) Then
                        High_Load = Volt_Schedule.Range("G1").Offset(kV_Row - 1, 3 * Load_Index + 1)
                        If Load >= Low_Load And Load <= High_Load Then
                            Ref_Voltage = Load_Voltage
                            Exit For
                        End If
                    Else
                        If Load >= Low_Load Then
                            Ref_Voltage = Load_Voltage
                            Exit For
                        End If
                    End If
                End If
            Next
            
            Low_V = 0.98 * Ref_Voltage
            High_V = 1.02 * Ref_Voltage
            .Range("M1").Offset(DataPoint, 0) = "[" & Low_V & ", " & High_V & "]"
            
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
    Dim Outside_1 As Long
    Dim Outside_1_5 As Long
    Dim Outside_2 As Long
    Dim num_data As Long
    Dim index As Integer
    Dim QLength As Long
    
    num_data = 0
    
    With ws
        For index = 0 To 5
            QLength = Quarter_Lengths(index)
            .Range("S15").Offset(index + 1, 1) = QLength
            If QLength <> 0 Then
                Outside_1 = Application.Sum(.Range("O3").Offset(num_data, 0).Resize(QLength))
                Outside_1_5 = Application.Sum(.Range("P3").Offset(num_data, 0).Resize(QLength))
                Outside_2 = Application.Sum(.Range("Q3").Offset(num_data, 0).Resize(QLength))
                .Range("S15").Offset(index + 1, 2) = Outside_1
                .Range("S15").Offset(index + 1, 3) = Format(Outside_1 / QLength, "Percent")
                .Range("S15").Offset(index + 1, 4) = Outside_1_5
                .Range("S15").Offset(index + 1, 5) = Format(Outside_1_5 / QLength, "Percent")
                .Range("S15").Offset(index + 1, 6) = Outside_2
                .Range("S15").Offset(index + 1, 7) = Format(Outside_2 / QLength, "Percent")
                num_data = num_data + QLength
            End If
        Next
    End With
End Sub
