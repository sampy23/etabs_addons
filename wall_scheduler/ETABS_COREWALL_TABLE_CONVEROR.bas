Attribute VB_Name = "ETABS_COREWALL_TABLE_CONVEROR"
' This program converts corewall data from etabs into user desired format using pivot table and custom made area convertor
                                        ' --------------------------------------------------- A VBA SCRIPT BY SHAHABAZ SHA

Public Function ArrayLen(arr As Variant) As Integer
    ' This function returns length of given array
    
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
Public Function Area_convertor(x As Long) As String
    ' This function converts the input area into T16@200 format
    
    Dim Area_matrix As Variant, corr_value As Variant
    Dim i As Integer, j As Integer, Area_matrix_length As Integer
    Set ASheet = Worksheets("Area Sheet")
    
    ' find maximum row in area sheet
    MaxRow_table = ASheet.Range("C" & Rows.Count).End(xlUp).Row
    Area_matrix = ASheet.Range("B2:C" & MaxRow_table).Value
    Area_matrix_length = ArrayLen(Area_matrix)
    corr_value = ASheet.Range("D2:D" & MaxRow_table)
        j = 1
        Do While j <= Area_matrix_length
            If x >= Area_matrix(j, 1) And x < Area_matrix(j, 2) Then
                Area_convertor = corr_value(j, 1) ' offset adds the row
            End If
            j = j + 1
        Loop
End Function

Sub Convertor()

Dim PTable As PivotTable
Dim PCache As PivotCache
Dim PRange As Range
Dim myRng       As Range
Dim mycell      As Range
Dim PSheet As Worksheet
Dim DSheet As Worksheet
Dim RSheet As Worksheet
Dim lastRow As Long
Dim lastColumn As Long
Dim rebar_dia As String
Dim rebar_spacing As String
Dim rebar_area As String

On Error Resume Next
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Worksheets("Pivot Sheet").Delete 'This will delete the exisiting pivot table worksheet
Worksheets.Add After:=ActiveSheet ' This will add new worksheet
ActiveSheet.Name = "Pivot Sheet" ' This will rename the worksheet as "Pivot Sheet"
On Error GoTo 0

Set ISheet = Worksheets("Input")
Set DSheet = Worksheets("Data Sheet")
Set PSheet = Worksheets("Pivot Sheet")
Set RSheet = Worksheets("Result Sheet")

RSheet.Cells.ClearContents
' name the column name
rebar_dia = "VL" 'VERTICAL
rebar_spacing = "SPACING"
rebar_area = "HZ" 'HORIZONTAL

'Find Last used row and column in data sheet
lastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row
lastColumn = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

'Set the pivot table data range
Set PRange = DSheet.Cells(1, 1).Resize(lastRow, lastColumn)

'Set pivot cahe
Set PCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, SourceData:=PRange)

'Create blank pivot table
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="Shear Wall Reinforcement")

'Insert country to Row Filed
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Story")
.Orientation = xlRowField
.Position = 1
End With

'Insert Product to Row Filed & position 2
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Station")
.Orientation = xlRowField
.Position = 2
End With

'Insert Segment to Column Filed & position 1
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Pier Label")
.Orientation = xlColumnField
.Position = 1
End With

'Insert Sales column to the data field
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Edge Rebar")
.Orientation = xlDataField
.Position = 1
.Function = xlMax
End With
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Rebar Spacing")
.Orientation = xlDataField
.Position = 2
.Function = xlMin
End With
With PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Shear Rebar")
.Orientation = xlDataField
.Position = 3
.Function = xlMax
End With

' filter out station field
PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Station").Orientation = xlHidden

'rename vlaue fields
PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Max of Edge Rebar").Caption = rebar_dia
PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Min of Rebar Spacing").Caption = rebar_spacing
PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Max of Shear Rebar").Caption = rebar_area

'Grand total off for Rows and Columns
PSheet.PivotTables("Shear Wall Reinforcement").ColumnGrand = False
PSheet.PivotTables("Shear Wall Reinforcement").RowGrand = False

'sorting the first column
PSheet.PivotTables("Shear Wall Reinforcement").PivotFields("Story") _
                        .AutoSort xlDescending, "Story"
'copy and paste special pivot table
PSheet.PivotTables("Shear Wall Reinforcement").TableRange2.Copy
With RSheet.Range("A1")
    .PasteSpecial Paste:=xlPasteValues
    .PasteSpecial Paste:=xlPasteFormats
    .PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
     SkipBlanks:=False, Transpose:=False
End With


' find last row and column of result sheet
lastRow = RSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
lastColumn = RSheet.Range("A1").CurrentRegion.Columns.Count

For i = 1 To lastColumn
        If Cells(3, i) = rebar_area Then ' search for column header in 3rd row
            'select all values in the column starting from 4th row
            Set myRng = RSheet.Range(RSheet.Cells(4, i), RSheet.Cells(lastRow, i))
            For Each mycell In myRng
                If IsEmpty(mycell) = False Then ' if cell is not empty
                    mycell.Value = Area_convertor(mycell.Value)
                Else
                     mycell.Value = "----"
                End If
            Next
        End If
Next

'This snippet concat two columns
Dim lRow As Long
Dim lCol As Long


lCol = RSheet.Cells(2, Columns.Count).End(xlToLeft).Column
lRow = RSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lCol Step 2
For j = 4 To lRow
If IsEmpty(RSheet.Cells(j, i)) = False Then
    RSheet.Cells(j, i) = "T" & RSheet.Cells(j, i) & "@" & RSheet.Cells(j, i + 1)
Else
    RSheet.Cells(j, i) = "----"
End If
    
Next j
i = i + 1
Next i

For i = 2 To lCol Step 2
RSheet.Columns(i + 1).EntireColumn.Delete
Next i


RSheet.Range(RSheet.Cells(1, 1), RSheet.Cells(lRow, lCol)).ClearFormats
RSheet.Range(RSheet.Cells(1, 1), RSheet.Cells(lRow, lCol)).Font.Name = "Arial"
RSheet.Range(RSheet.Cells(1, 1), RSheet.Cells(lRow, lCol)).Font.Size = 12
RSheet.Range(RSheet.Cells(1, 1), RSheet.Cells(lRow, lCol)).Columns.AutoFit

ISheet.Cells.ClearContents
End Sub





