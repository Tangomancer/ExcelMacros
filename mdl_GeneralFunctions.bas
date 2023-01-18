Attribute VB_Name = "mdl_GeneralFunctions"
Option Explicit


Function ConcatRange(RowRange As Range, Optional strSeparator As String) As String
    Dim x As Long, CellVal As String, ReturnVal As String, Result As String, Delimiter As String
    
    If strSeparator = "" Then
        Delimiter = ", "
    Else
        Delimiter = strSeparator
    End If
    
    For x = 1 To RowRange.Count
        ReturnVal = RowRange(x).Value
        If Len(RowRange(x).Value) Then
            If InStr(Result & Delimiter, Delimiter & ReturnVal & Delimiter) = 0 Then
                Result = Result & Delimiter & ReturnVal
            End If
        End If
    Next

    ConcatRange = Mid(Result, Len(Delimiter) + 1)

End Function





Sub SimplePivotLayout()
    On Error GoTo Err_SimplePivotLayout
    Dim err_count As Integer
    err_count = 0
    
    Dim oldStatusbar As Boolean
    Dim PT As PivotTable
    Dim PTF As PivotField
    
    Dim PivotName As String
    
    Set PT = ActiveCell.PivotTable
    
    If Not PT Is Nothing Then
        oldStatusbar = Application.DisplayStatusBar
        Application.DisplayStatusBar = True
        Application.StatusBar = "Applying simple layout to pivot table"
    
    
        Application.ScreenUpdating = False
        With PT
            .InGridDropZones = True
            .RowAxisLayout xlTabularRow
    
            For Each PTF In .PivotFields
                'If Not PTF.SourceCaption = "" Then
                '    Application.StatusBar = "Applying simple layout to pivot field " & PTF.SourceCaption
                'Else
                    Application.StatusBar = "Applying simple layout to pivot field " & PTF.SourceName
                'End If
                PTF.Subtotals(1) = False
'Debug.Print PTF.Name
            Next PTF
        End With
        
        Application.StatusBar = False
        Application.DisplayStatusBar = oldStatusbar
        
    End If

Exit_SimplePivotLayout:
    Application.StatusBar = oldStatusbar
    Application.ScreenUpdating = True
    Set PT = Nothing
    Exit Sub

Err_SimplePivotLayout:
    err_count = err_count + 1
    If ((Err.Number = 1004) And (err_count < 2)) Then Resume Next 'Er zitten velden in de pivot waarvoor subtotals niet bepaald kunnen worden. Dan overslaan
    Application.ScreenUpdating = True
    MsgBox "Er gaat iets fout " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = False
    Resume Exit_SimplePivotLayout
        
End Sub


Sub SheetNames()
    Dim i As Long

    Columns(1).Insert
    For i = 1 To Sheets.Count
        Cells(i, 1) = Sheets(i).Name
    Next i
End Sub


Function LongestCommonSubstring(S1 As String, S2 As String) As String
  Dim MaxSubstrStart As Integer
  Dim MaxLenFound As Integer
  Dim i1, i2, x As Integer
  
  MaxSubstrStart = 1
  MaxLenFound = 0
  For i1 = 1 To Len(S1)
    For i2 = 1 To Len(S2)
      x = 0
      While i1 + x <= Len(S1) And _
            i2 + x <= Len(S2) And _
            Mid(S1, i1 + x, 1) = Mid(S2, i2 + x, 1)
        x = x + 1
      Wend
      If x > MaxLenFound Then
        MaxLenFound = x
        MaxSubstrStart = i1
      End If
    Next
  Next
  LongestCommonSubstring = Mid(S1, MaxSubstrStart, MaxLenFound)
End Function


Public Function LTrimZeros(CellValue)

     Dim strChr As String
     Dim n As Integer
     
     strChr = "0"
     n = 0
     
     If Not IsNull(CellValue) Then
         Do Until strChr <> "0"
             n = n + 1
             LTrimZeros = Mid(CellValue, n)
             strChr = Mid(CellValue, n, 1)
         Loop
     End If
         
 End Function


Sub FormatTimestamp()
    Selection.NumberFormat = "yyyy/mm/dd hh:mm:ss"
End Sub


Sub NumberToText()
On Error GoTo Err_NumberToText

    Dim rCell As Range
    Dim t, rCount As Long
    Dim s As Variant            'temporary cell value
    Dim strNrLength As Long     'desired length of string
    Dim strTrimLength As Long   'lengthh if original cell content exceeds strNrLength
    Dim oldStatusbar As Boolean
    
    'Ask length of number including leading zero's
    strNrLength = InputBox("Gewenste lengte van nummer met voorloopnullen?", "Voer in")
    
    
    rCount = Selection.Cells.Count
    t = 0
    oldStatusbar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.StatusBar = "Nrs completed: " & t
    
    Application.Calculation = xlCalculationManual
    
    Application.ScreenUpdating = False
    
    For Each rCell In Selection
        t = t + 1
        
        If Not rCell.HasFormula Then
            'first remove leading 0, necessary for values already longer than strNrLength including the leading 0
            s = rCell.Value
            While Left(s, 1) = "0" And s <> ""
                s = Mid(s, 2)
            Wend
            'set result length to the maximum of the actual length of s (without leading 0) or strNrLength
            If Len(s) > strNrLength Then
                strTrimLength = Len(s)
            Else
                strTrimLength = strNrLength
            End If
            rCell.NumberFormat = "@"                                     'sets cell format as text
            rCell = Right(String(strNrLength, "0") & rCell.Value, strTrimLength)
        End If
        
        If t Mod 1000 = 0 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "Nrs completed: " & Format(t, "#,##0") & " of " & Format(rCount, "#,##0")
            Application.ScreenUpdating = False
        End If
        
    Next rCell
    
    Application.Calculation = xlCalculationAutomatic
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusbar

   
Exit_NumberToText:
    
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = oldStatusbar
    Application.ScreenUpdating = True
    Exit Sub

Err_NumberToText:
    Application.ScreenUpdating = True
    MsgBox "Er gaat iets fout " & Err.Number & ": " & Err.Description
    Application.ScreenUpdating = False
    Resume Exit_NumberToText
    
End Sub


Sub HighlightStrings()
    Dim xHStr As String, xStrTmp As String
    Dim xHInt As Integer
    Dim xHStrLen As Long, xCount As Long, I As Long
    Dim xCell As Range
    Dim xArr
    On Error Resume Next
    xHStr = Application.InputBox("What is the string to highlight:", "Please enter...", , , , , , 2)
    xHInt = Application.InputBox("What is the color to use (e.g: 3=red, 4=green, 5=blue, 6=yellow, 7=magenta):", "Please enter...", 3, , , , , 1)
    If TypeName(xHStr) <> "String" Then Exit Sub
    Application.ScreenUpdating = False
        xHStrLen = Len(xHStr)
        For Each xCell In Selection
            xArr = Split(xCell.Value, xHStr)
            xCount = UBound(xArr)
            If xCount > 0 Then
                xStrTmp = ""
                For I = 0 To xCount - 1
                    xStrTmp = xStrTmp & xArr(I)
                    xCell.Characters(Len(xStrTmp) + 1, xHStrLen).Font.ColorIndex = xHInt
                    xStrTmp = xStrTmp & xHStr
                Next
            End If
        Next
    Application.ScreenUpdating = True
End Sub


Sub ResizeAllChartObjects()
'Apply Activechart sizes for both chart and plotarea to all
'other charts on this page.
'
    Dim objDefaultChart As Chart
    Dim objChart As ChartObject
    Dim intIndex As Integer
     
    On Error Resume Next
    If ActiveSheet.ChartObjects.Count > 1 Then
        ' only bother if there are more than 1 chartobject
        Set objDefaultChart = ActiveChart
        If Not (objDefaultChart Is Nothing) Then
            For intIndex = 1 To ActiveSheet.ChartObjects.Count
                Set objChart = ActiveSheet.ChartObjects(intIndex)
                If objChart.Name = objDefaultChart.Name Then
                    ' This chart is already correct
                Else
                    With objChart
                        .Width = objDefaultChart.Parent.Width
                        .Height = objDefaultChart.Parent.Height
                        With .Chart.PlotArea
                            .Width = objDefaultChart.PlotArea.Width
                            .Height = objDefaultChart.PlotArea.Height
                            .Left = (objDefaultChart.Parent.Width - objDefaultChart.PlotArea.Width) / 2
                            .Top = (objDefaultChart.Parent.Height - objDefaultChart.PlotArea.Height) / 2
                        End With
                    End With
                End If
            Next
        Else
            ' No Active chart
            MsgBox "Please select chart on which to base sizes", vbExclamation
        End If
    End If
End Sub
