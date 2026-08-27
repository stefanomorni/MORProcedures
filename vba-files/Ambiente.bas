Attribute VB_Name = "Ambiente"
Sub CancelNamedQueryAndWait(Optional ByVal QueryName As String = "")
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim co As WorkbookConnection
    Dim stillRefreshing As Boolean
    Dim found As Boolean
    
    ' --- Exit if no query name provided ---
    If Trim(QueryName) = "" Then Exit Sub
    
    ' Target active workbook (not the add-in)
    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    If wb Is ThisWorkbook Then Exit Sub ' avoid running on the add-in
    
    ' --- Try to find and cancel a QueryTable with that name ---
    found = False
    For Each ws In wb.Worksheets
        For Each qt In ws.QueryTables
            If StrComp(qt.Name, QueryName, vbTextCompare) = 0 Then
                found = True
                If qt.Refreshing Then qt.CancelRefresh
                ' Wait until it stops refreshing
                Do While qt.Refreshing
                    DoEvents
                Loop
                Exit Sub ' We’re done
            End If
        Next qt
    Next ws
    
    ' --- Try to find and cancel a WorkbookConnection with that name ---
    If Not found Then
        For Each co In wb.Connections
            If StrComp(co.Name, QueryName, vbTextCompare) = 0 Then
                found = True
                On Error Resume Next
                If co.Refreshing Then co.CancelRefresh
                On Error GoTo 0
                ' Wait until it stops refreshing
                Do
                    stillRefreshing = False
                    On Error Resume Next
                    stillRefreshing = co.Refreshing
                    On Error GoTo 0
                    If stillRefreshing Then DoEvents
                Loop While stillRefreshing
                Exit Sub ' Done
            End If
        Next co
    End If
    
    ' --- If not found, do nothing ---
End Sub


