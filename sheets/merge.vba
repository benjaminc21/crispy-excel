Public Function last_row(sheet_idx) As Integer
    With Sheets(sheet_idx)
        If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
            lastrow = .Cells.Find(What:="*", _
                After:=Range("A1"), _
                Lookat:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
        Else
            lastrow = 1
        End If
        last_row = lastrow
    End With
End Function

Public Function last_col(sheet_idx) As Integer
    With Sheets(sheet_idx)
            If Application.WorksheetFunction.CountA(.Cells) <> 0 Then
                lastcol = .Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    Lookat:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Column
            Else
                lastcol = 1
            End If
            last_col = lastcol
        End With
End Function

Sub Macro1()
'
' Macro1 Macro
'
 
' merge multiple selected sheets into one master sheet
 
On Error Resume Next
 
' always ensure that the last sheet is named 'combined'
output_sheet = "combined"

Dim sheet_idx As Integer

Dim combined_last_row As Integer
Dim combined_last_col As Integer

combined_last_row = 1

For sheet_idx = 1 To Sheets.Count - 1
    
    
    
    Dim source_sheet_last_col As Integer
    Dim dest_sheet_last_col As Integer
    
    Sheets(sheet_idx).Activate
    source_sheet_last_col = last_col(sheet_idx)
        
    Selection.CurrentRegion.Select
    Selection.Copy
    
    Sheets("combined").Activate
    dest_sheet_last_col = last_col(Sheets.Count)
    
    If sheet_idx <> 1 Then
        If dest_sheet_last_col <> source_sheet_last_col Then
            MsgBox ("Data Integrity Warning: There may be a data mismatch.")
        End If
    End If
    
    Cells(combined_last_row, 1).Select
    ActiveSheet.Paste
    
    combined_last_row = last_row(sheet_idx) + combined_last_row
Next

End Sub

