Option Explicit

Sub MigrateData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim new_wb As Workbook
    Dim new_ws As Worksheet
    Dim savePath As String
    Dim wbPath As String
    Dim fileName As String
    Dim new_ws_name As String
    Dim counter As Integer
    
    wbPath = "C:\Users\chanjc\Desktop\Volt_VAR Analysis\"
    fileName = Dir(wbPath)
    savePath = "C:\Users\chanjc\Desktop\"
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    counter = 0
    Do While fileName <> ""
        Workbooks.Open (wbPath & fileName)
        For Each ws In Workbooks(fileName).Worksheets
            If ws.Name <> "VAR Schedules" And ws.Name <> "Volt Schedules" Then
                Set new_wb = Workbooks.Add
                new_ws_name = ws.Name
                new_wb.Worksheets.Add().Name = new_ws_name
                Set new_ws = new_wb.Worksheets(new_ws_name)

                'Boiler plate
                new_ws.Range("A1") = "Date"
                new_ws.Range("B1") = "Time"
                new_ws.Range("C1") = "MW"
                new_ws.Range("D1") = "MVAR"
                new_ws.Range("E1") = "In Range"
                new_ws.Range("F1") = "> -10 MVAR"
                new_ws.Range("G1") = "> -20 MVAR"
                new_ws.Range("H1") = "<= -20 MVAR"
                new_ws.Range("I1") = "Upper VAR Threshold"
                new_ws.Range("J1") = "Lower VAR Threshold"
                new_ws.Range("K1") = "Actual Difference"
                new_ws.Range("L1") = "kV"
                new_ws.Range("M1") = "Reference Voltage"
                new_ws.Range("N1") = "Difference (kV)"
                new_ws.Range("O1") = "Outside 1%"
                new_ws.Range("P1") = "Outside 1.5%"
                new_ws.Range("Q1") = "Outside 2%"

                'Copy over the date and time
                ws.Range("A3:B656582").Copy Destination:=new_ws.Range("A2:B656581")
                'Copy over the MW load
                ws.Range("D3:D656582").Copy Destination:=new_ws.Range("C2:C656581")
                'Copy over the MVAR load
                ws.Range("I3:I656582").Copy Destination:=new_ws.Range("D2:D656581")
                'Copy over the kV
                ws.Range("W3:W656582").Copy Destination:=new_ws.Range("L2:L656581")

                new_wb.SaveAs fileName:=savePath & counter & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
                new_wb.Close True
                Workbooks(fileName).Close
            End If
        Next ws
        counter = counter + 1
        fileName = Dir()
    Loop
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


