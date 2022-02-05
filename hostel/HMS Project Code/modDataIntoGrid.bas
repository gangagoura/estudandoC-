Attribute VB_Name = "modDataIntoGrid"
Option Explicit

Private Sub HighlightGridRow(grd As MSFlexGrid, iRow As Long)
    With grd
        If .Rows > 1 Then
            .Row = iRow
            .Col = 1
            .ColSel = .Cols - 1
            .RowSel = iRow
        End If
    End With
End Sub
Public Sub LoadRecordsetIntoGrid(rs As Recordset, grd As MSFlexGrid, _
                Optional AutosizeColumns As Boolean = True, _
                Optional HighlightFirstRow As Boolean = False)
                
    Dim x As Long
    Dim Count As Long

    grd.Redraw = False
    grd.Clear

    'setup the minimum number of rows & add column headers
    grd.Rows = 2
    grd.FixedRows = 1
    grd.Row = 0
    grd.Cols = rs.Fields.Count + 1
    For x = 0 To rs.Fields.Count - 1
        grd.Col = x + 1
        grd.Text = rs.Fields(x).Name
        grd.ColData(x + 1) = rs.Fields(x).Type
    Next

    If rs.CursorLocation = adUseClient Then
        grd.Rows = rs.RecordCount + 1
        For Count = 1 To rs.RecordCount
            
            grd.TextMatrix(Count, 0) = Count    'assign line number
            For x = 0 To rs.Fields.Count - 1
                'we use Variant conversion to avoid any possible NULL errors
                grd.TextMatrix(Count, x + 1) = "" & CVar(rs.Fields(x).Value)
            Next
            rs.MoveNext
        Next
    ElseIf rs.CursorLocation = adUseServer Then
        Do While Not rs.EOF
            Count = Count + 1
            
            If Count >= grd.Rows Then
                'increase the amount of rows by 100 at a time - makes everything faster
                grd.Rows = grd.Rows + 100
            End If
            
            grd.TextMatrix(Count, 0) = Count    'assign line number
            For x = 0 To rs.Fields.Count - 1
                'we use Variant conversion to avoid any possible NULL errors
                grd.TextMatrix(Count, x + 1) = "" & CVar(rs.Fields(x).Value)
            Next
    
            rs.MoveNext
        Loop
        grd.Rows = Count + 1    'reset the number of rows to the number of records
    End If
    
    'autosize the grid
    If AutosizeColumns Then SetGridColumnWidth grd

    'highlight the first row
    If HighlightFirstRow Then
        If grd.Rows > 1 Then HighlightGridRow grd, 1
    End If
    
    'redraw the entire grid
    grd.Redraw = True
End Sub
Public Sub LoadRecordsetIntoGridNoCr(rs As Recordset, grd As MSFlexGrid, _
                Optional AutosizeColumns As Boolean = True, _
                Optional HighlightFirstRow As Boolean = True)
                
    'This sub assumes that the records doesn't have any
    'data in it with Carriage Returns
    'This knowledge allows the routine to dump the
    'entire recordset wholesale into the MsFlexGrid via the .Clip property
    
    Dim x As Long
    Dim Count As Long

    'this prevents the grid from repainting everytime something is added
    grd.Redraw = False
    grd.Clear

    'setup the minimum number of rows & add column headers
    grd.Rows = 2
    grd.FixedRows = 1
    grd.Row = 0
    grd.Cols = rs.Fields.Count + 1
    For x = 0 To rs.Fields.Count - 1
        grd.Col = x + 1
        grd.Text = rs.Fields(x).Name
        grd.ColData(x + 1) = rs.Fields(x).Type
    Next

    If rs.CursorLocation = adUseClient Then
        If rs.RecordCount > 0 Then
            'setup all the rows necessary for this recordset
            grd.Rows = rs.RecordCount + 1
            grd.Row = 1
            grd.Col = 1
                        
            grd.ColSel = rs.Fields.Count
            grd.RowSel = rs.RecordCount
            
            'fill in the recordset
            grd.Clip = rs.GetString(, , Chr$(9), Chr$(13))
            
            'remove the highlight
            grd.RowSel = 0
            grd.ColSel = 0
            grd.Row = 1
            grd.Col = 1
            
            'assign record numbers
            For x = 1 To grd.Rows - 1
                grd.TextMatrix(x, 0) = x
            Next
            
        End If
    ElseIf rs.CursorLocation = adUseServer Then
        Do While Not rs.EOF
            Count = Count + 1
            
            If Count >= grd.Rows Then
                'increase the amount of rows by 100 at a time - makes everything faster
                grd.Rows = grd.Rows + 100
            End If
            
            grd.TextMatrix(Count, 0) = Count    'assign line number
            For x = 0 To rs.Fields.Count - 1
                'we use Variant conversion to avoid any possible NULL errors
                grd.TextMatrix(Count, x + 1) = "" & CVar(rs.Fields(x).Value)
            Next
    
            rs.MoveNext
        Loop
        grd.Rows = Count + 1    'reset the number of rows to the number of records
    End If
    
    'autosize the grid
    If AutosizeColumns Then SetGridColumnWidth grd

    'highlight the first row
    If HighlightFirstRow Then
        If grd.Rows > 1 Then HighlightGridRow grd, 1
    End If
    
    'redraw the entire grid
    grd.Redraw = True
End Sub
Private Sub SetGridColumnWidth(grd As MSFlexGrid)
    'params:    ms flexgrid control
    'purpose:   sets the column widths to the lengths of the longest string in the column
    'requirements:  the grid must have the same font as the underlying form

    Dim InnerLoopCount As Long
    Dim OuterLoopCount As Long
    Dim lngLongestLen As Long
    Dim sLongestString As String
    Dim lngColWidth As Long
    Dim szCellText As String

    For OuterLoopCount = 0 To grd.Cols - 1
        sLongestString = ""
        lngLongestLen = 0

        'grd.Col = OuterLoopCount
        For InnerLoopCount = 0 To grd.Rows - 1
            szCellText = grd.TextMatrix(InnerLoopCount, OuterLoopCount)
            'grd.Row = InnerLoopCount
            'szCellText = Trim$(grd.Text)
            If Len(szCellText) > lngLongestLen Then
                lngLongestLen = Len(szCellText)
                sLongestString = szCellText
            End If
        Next
        lngColWidth = grd.Parent.TextWidth(sLongestString)
        grd.ColWidth(OuterLoopCount) = lngColWidth + 200    'add 100 for more readable spreadsheet
    Next
End Sub





