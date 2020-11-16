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


