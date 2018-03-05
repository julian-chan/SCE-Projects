Attribute VB_Name = "Data_Cleanup"
Sub Main()
    Call RemoveExcessData
    Call RemoveOutliers
End Sub

Sub RemoveExcessData()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = Worksheets("tmp0006")

    With ws
        For i = 0 To 60
            .Range("A787623").Offset(i, 0).ClearContents
            .Range("B787623").Offset(i, 0).ClearContents
            .Range("C787623").Offset(i, 0).ClearContents
            .Range("D787623").Offset(i, 0).ClearContents
        Next i
    End With
End Sub

Sub RemoveOutliers()
    Dim ws As Worksheet
    Dim i As Long
    Dim iter As Long
    Dim average As Double
    Dim before As Long
    Dim after As Long
    Dim outlier As Boolean
    Set ws = Worksheets("tmp0006")
    
    outlier = False
    
    With ws
        For i = 1 To 787620
            If .Range("L2").Offset(i, 0) < 501 And Not outlier Then
                outlier = True
                before = i - 1
            End If
            
            If .Range("L2").Offset(i, 0) > 501 And outlier Then
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


