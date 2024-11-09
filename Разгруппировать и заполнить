Sub UnMerge_And_Fill_By_Value() ' разгруппировать все ячейки в Selection и ячейки каждой бывшей группы заполнить значениями из их первых ячеек
     Dim Address As String
     Dim Cell As Range
 
     If TypeName(Selection) <> "Range" Then
         Exit Sub
     End If
 
     If Selection.Cells.Count = 1 Then
         Exit Sub
     End If
 
     Application.ScreenUpdating = False
 
     For Each Cell In Intersect(Selection, ActiveSheet.UsedRange).Cells
         If Cell.MergeCells Then
             Address = Cell.MergeArea.Address
             Cell.Unmerge
             Range(Address).Value = Cell.Value
         End If
     Next
End Sub

Sub MergeCls()
 Dim ri As Integer, r2 As Integer, Col As Integer
 r1 = ActiveCell.Row
 r2 = ActiveCell.Row
 Col = ActiveCell.Column
 Do
 If Cells(r1, Col) <> Cells(r2 + 1, Col) Then
 If r1 <> r2 Then
 Range(Cells(r1 + 1, Col), Cells(r2, Col)).ClearContents
 With Range(Cells(r1, Col), Cells(r2, Col))
 .HorizontalAlignment = xlCenter
 .VerticalAlignment = xlCenter
 .WrapText = True
 .Orientation = 0
 .AddIndent = False
 .IndentLevel = 0
 .ShrinkToFit = False
 .ReadingOrder = xlContext
 .MergeCells = True
 End With
 End If
 r1 = r2 + 1
 End If
 r2 = r2 + 1
 Loop Until Cells(r2, Col) = ""
 End Sub
