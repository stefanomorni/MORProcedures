Attribute VB_Name = "UserInteraction"
'@Folder("MORProcedures.Modules")
Option Explicit
Public RisultatoFinestra                         'Valore restituito dalla finestra

Sub MessaggioInBarraDiStato(Messaggio As String, DurataInSecondi As Integer)

    'Displays the "Messaggio" in the Status bar for "DurataInSecondi" seconds

    'Declarations

    Dim oldStatusBar As String
    Dim UpdateTime As Date

    'Procedure

    With Application
        ResetBarraDiStato
        .StatusBar = Messaggio
        UpdateTime = Now() + TimeValue("00:" & Int(DurataInSecondi / 60) & ":" & DurataInSecondi - Int(DurataInSecondi / 60) * 60)
        If DurataInSecondi > 0 Then Application.OnTime EarliestTime:=UpdateTime, procedure:="ResetBarraDiStato"
    End With
End Sub

Sub ResetBarraDiStato()
    Application.StatusBar = False
End Sub

Sub FinestraInformativa(Titolo As String, Messaggio As String, Modeless As Boolean)
    With FrmFinestraInformativa
        .Caption = Titolo
        .Messaggio.value = Messaggio
        If Modeless Then
            .CbxModeless.value = True
        Else
            .CbxModeless.value = False
        End If
        .Show
    End With
End Sub

Sub Finestra2Alternative(Titolo As String, Messaggio As String, Modeless As Boolean, CaptionAlternativa_1 As String, CaptionAlternativa_2 As String)


    With Frm2Alternative
        .Caption = Titolo
        .Messaggio.value = Messaggio
        .CmdAlternativa1.Caption = CaptionAlternativa_1
        .CmdAlternativa2.Caption = CaptionAlternativa_2
        If Modeless Then
            .CbxModeless.value = True
        Else
            .CbxModeless.value = False
        End If
        .Show
    End With
End Sub

Sub Finestra1Range(Titolo As String, Messaggio As String, Modeless As Boolean)

    With Frm1Range
        .Caption = Titolo
        .MessaggioRange = Messaggio
        .Riferimento = Selection.AddressLocal
        If Modeless Then
            .CbxModeless.value = True
        Else
            .CbxModeless.value = False
        End If
        .Show
    End With
End Sub

Sub Highlight_Selected_Tabe_Row(Table As Object, Selezione As Object)
    ' Ignore whole-row or large multi-cell selections (e.g. during row deletion)
    If Selezione Is Nothing Or Table Is Nothing Then Exit Sub
    If Selezione.Cells.CountLarge > 50 Then Exit Sub

    On Error GoTo CleanExit

    Dim ws As Worksheet
    Set ws = Table.Parent

    ' Only run when the selection intersects the table body
    If Intersect(Selezione, Table.DataBodyRange) Is Nothing Then Exit Sub

    ' 1. Self-healing check: ensure the sheet-scoped ActiveTableRow name exists
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names("ActiveTableRow")
    On Error GoTo CleanExit

    If n Is Nothing Then
        ws.Names.Add Name:="ActiveTableRow", RefersTo:="=" & Selezione.Row
        EnsureTableRowHighlightRule Table
    Else
        ' 2. Ultra-fast update: updates native crosshair in 0ms without touching cells
        n.RefersTo = "=" & Selezione.Row
    End If

CleanExit:
    On Error GoTo 0
End Sub

Private Sub EnsureTableRowHighlightRule(Table As Object)
    On Error Resume Next
    Dim tblRange As Range
    Set tblRange = Table.DataBodyRange
    If tblRange Is Nothing Then Exit Sub

    ' Check if rule already exists to avoid duplication
    Dim i As Long, hasRule As Boolean
    For i = 1 To tblRange.FormatConditions.count
        If InStr(1, tblRange.FormatConditions(i).Formula1, "ActiveTableRow", vbTextCompare) > 0 Then
            hasRule = True
            Exit For
        End If
    Next i

    ' If missing, add the native conditional formatting rule
    If Not hasRule Then
        Dim fc As FormatCondition
        Set fc = tblRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=ROW()=ActiveTableRow")
        With fc.Interior
            .Color = 10092543 ' Pale yellow
            .Pattern = xlSolid
        End With
    End If
    On Error GoTo 0
End Sub
