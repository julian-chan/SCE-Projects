Attribute VB_Name = "Module2"
Sub CopySchedules()
    Dim var As Worksheet
    Dim volt As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wbPath As String
    Dim fileName As String
    Dim contains_var As Boolean
    Dim contains_volt As Boolean
    Dim contains_sheet As Boolean
    
    Application.DisplayAlerts = False
    
    wbPath = "C:\Users\chanjc\Desktop\"
    fileName = Dir(wbPath)
    
    Workbooks.Open (wbPath & "Volt_VAR Analysis\References\VAR Schedules.xlsx")
    Set var = Workbooks("VAR Schedules.xlsx").Worksheets("VAR Schedules")
    Workbooks.Open (wbPath & "Volt_VAR Analysis\References\Volt Schedules.xlsx")
    Set volt = Workbooks("Volt Schedules.xlsx").Worksheets("Volt Schedules")
    
    contains_var = False
    contains_volt = False
    contains_sheet = False
    
    Do While fileName <> ""
        If InStr(fileName, ".xlsm") And fileName <> "Dummy.xlsm" Then
            Workbooks.Open (wbPath & fileName)
            Set wb = Workbooks(fileName)
            
            For Each ws In Workbooks(fileName).Worksheets
                If ws.Name = "VAR Schedules" Then
                    contains_var = True
                ElseIf ws.Name = "Volt Schedules" Then
                    contains_volt = True
                ElseIf Left(ws.Name, 5) = "Sheet" Then
                    contains_sheet = True
                End If
            Next ws
            
            If Not contains_var Then
                var.Copy Before:=wb.Worksheets(Left(wb.Name, Len(wb.Name) - 5))
            End If
            
            If Not contains_volt Then
                volt.Copy Before:=wb.Worksheets(Left(wb.Name, Len(wb.Name) - 5))
            End If
            
            If contains_sheet Then
                Worksheets("Sheet1").Delete
            End If

            wb.Close True
        End If
        
        contains_var = False
        contains_volt = False
        contains_sheet = False
        fileName = Dir()
    Loop

End Sub
