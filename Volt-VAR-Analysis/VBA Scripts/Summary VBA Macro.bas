Option Explicit
Option Compare Text

Sub GetSummary()
    Dim directory As String
    Dim fileName As String
    Dim sheet As Worksheet
    Dim summary As Worksheet
    Dim substation As String
    Dim index As Integer
    Dim quarter As Integer
    Dim out1 As Double
    Dim out1_5 As Double
    Dim out2 As Double
    
    Set summary = Worksheets("Voltage Summary")
    
    directory = Application.ActiveWorkbook.path & "\"
    fileName = Dir(directory & "*.xlsm")
    
    Do While fileName <> ""
        If fileName <> "Volt Summary.xlsm" Then
            Workbooks.Open (directory & fileName)
            For Each sheet In Workbooks(fileName).Worksheets
                If sheet.Name <> "VAR Schedules" And sheet.Name <> "Volt Schedules" Then
                    substation = sheet.Name
                    
                    For index = 0 To 42
                        If substation = summary.Range("C6").offset(index, 0) Then
                            For quarter = 0 To 4
                                out1 = sheet.Range("A656595").offset(quarter + 1, 3)
                                out1_5 = sheet.Range("A656595").offset(quarter + 1, 5)
                                out2 = sheet.Range("A656595").offset(quarter + 1, 7)
                                
                                summary.Range("C6").offset(index, 4 * quarter + 2) = Format(out1, "Percent")
                                summary.Range("C6").offset(index, 4 * quarter + 3) = Format(out1_5, "Percent")
                                summary.Range("C6").offset(index, 4 * quarter + 4) = Format(out2, "Percent")
                                
                                If out2 > 0.2 Then
                                    summary.Range("C6").offset(index, 4 * quarter + 4).Interior.Color = RGB(255, 102, 102)
                                    summary.Range("C6").offset(index, 4 * quarter + 4).Font.Color = RGB(139, 0, 0)
                                End If
                            Next
                            Exit For
                        End If
                    Next
                End If
            Next sheet
            Workbooks(fileName).Close
        End If
        fileName = Dir()
    Loop
    
End Sub

