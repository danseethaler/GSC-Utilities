Attribute VB_Name = "ContextMenus"
Option Explicit

Sub PasteAndDeliminateComma(control As IRibbonControl)

Application.ScreenUpdating = False

    ActiveSheet.Paste
        
On Error Resume Next

    Selection.TextToColumns _
      Destination:=ActiveCell, _
      DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=False, _
      Semicolon:=False, _
      Comma:=True, _
      Space:=False, _
      OtherChar:=""
      
      
Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True

Application.ScreenUpdating = True

End Sub

Sub PasteAndDeliminateSpace(control As IRibbonControl)

Application.ScreenUpdating = False

    ActiveSheet.Paste
        
On Error Resume Next
        
    Selection.TextToColumns _
      Destination:=ActiveCell, _
      DataType:=xlDelimited, _
      TextQualifier:=xlDoubleQuote, _
      ConsecutiveDelimiter:=False, _
      Tab:=True, _
      Semicolon:=False, _
      Comma:=False, _
      Space:=True, _
      Other:=False

Application.ScreenUpdating = True

End Sub

Sub ConcatenateDelimitedText(control As IRibbonControl) '
Dim ConcatRange As Range
Dim cell As Range

Set ConcatRange = Intersect(Selection, Columns(ActiveCell.Column))

For Each cell In ConcatRange

Select Case Selection.Columns.Count
    Case 11: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8) & " " & cell.Offset(0, 9) & " " & cell.Offset(0, 10)
    Case 10: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8) & " " & cell.Offset(0, 9)
    Case 9: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7) & " " & cell.Offset(0, 8)
    Case 8: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6) & " " & cell.Offset(0, 7)
    Case 7: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5) & " " & cell.Offset(0, 6)
    Case 6: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4) & " " & cell.Offset(0, 5)
    Case 5: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3) & " " & cell.Offset(0, 4)
    Case 4: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2) & " " & cell.Offset(0, 3)
    Case 3: cell.Value = cell & " " & cell.Offset(0, 1) & " " & cell.Offset(0, 2)
    Case 2: cell.Value = cell & " " & cell.Offset(0, 1)

    Case Else: MsgBox ("Only 9 Columns can be concatenated with this function."): Exit Sub

End Select

Range(cell.Offset(0, 1), cell.Offset(0, Selection.Columns.Count - 1)).Clear

    Do Until Right(cell, 1) <> " "
        cell = Left(cell, Len(cell) - 1)
    Loop

Next

End Sub

