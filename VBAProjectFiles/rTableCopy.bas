Attribute VB_Name = "rTableCopy"
''Modify constants to indicate source & destination sheets & tables
Const SRCSHEET = "tImport"
Const SRCTABLE = "tblImport"
Const DSTSHEET = "tData"
Const DSTTABLE = "tblData"
                  
'' Copy all rows from source table to all rows (matching columns) in destination table
Sub Table2Table()
    ' This is useful for copying data from a Power Query table to a source table for analysis such as pivots
    ' 01.20.19 rmh Add some detective messages for Immediate window in case a sheet or table is not found
    
    dbg = True
    sheetFound = False
    tableFound = False
    
    Debug.Print ("== Table2Table " & Time())

    'Ensure requested sheets and tables exist
    If Not (ConfirmParms(SRCSHEET, SRCTABLE) And ConfirmParms(DSTSHEET, DSTTABLE)) Then
        If dbg Then Debug.Print ("** Sheet or table missing")
        MsgBox ("** Sheet or table missing. Check names.")
        Exit Sub
    Else
        If dbg Then Debug.Print (".. Sheet & Table found")
    End If
    
    'Prepare destination by removing all but one row
    Call DstTablePrep(DSTSHEET, DSTTABLE)
    'Perform copy
    Call TblSrc2Dst(SRCSHEET, SRCTABLE, DSTSHEET, DSTTABLE)
    
End Sub

'' Perform the copy of all table data to all table data
Private Sub TblSrc2Dst(Ssheet, Stable, Dsheet, Dtable)
    
    Dim st As ListObject
    Dim dt As ListObject
    
    Set st = Sheets(Ssheet).ListObjects(Stable)
    Set dt = Sheets(Dsheet).ListObjects(Dtable)
    
    'Basically...just a range copy/paste
    st.DataBodyRange.Copy Destination:=dt.DataBodyRange

End Sub

'' Remove all but first row of table
Private Sub DstTablePrep(Dsheet, Dtable)
    Dim ds As Sheets
    Dim dt As ListObject

    Set dt = Sheets(Dsheet).ListObjects(Dtable)
    'no real error checking for empty table, etc.
    On Error Resume Next
    dt.DataBodyRange.Offset(1, 0).Resize(dt.DataBodyRange.Rows.Count - 1, _
        dt.DataBodyRange.Columns.Count).Rows.Delete
    On Error GoTo 0
    
End Sub

''Ensure sheet and table exists
Private Function ConfirmParms(Qsheet, Qtable)

    Dim tbl As ListObject
    
    dbg = True
    ConfirmParms = False
    
    If IsSheet(Qsheet) Then
        'check table
        With Sheets(Qsheet)
            For Each tbl In .ListObjects
                If tbl.Name = Qtable Then
                    ConfirmParms = True
                    If dbg Then Debug.Print (".. Table(" & Qtable & ") exists? True")
                End If
            Next tbl
        End With
    End If
    If ConfirmParms = False Then
        Debug.Print ("** Unable to find " & Qtable & " on sheet " & Qsheet & ".")
    End If
    
End Function

'' True if a sheetname exists
Private Function IsSheet(Qsheet) As Boolean
    'Src: https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
    Dim sht As Worksheet
    Dim wb As Workbook
    
    dbg = True
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(Qsheet)
    On Error GoTo 0
    IsSheet = Not sht Is Nothing
    
    If dbg Then Debug.Print (".. Sheet(" & Qsheet & ") exists? " & IsSheet)

End Function

