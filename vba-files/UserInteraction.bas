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

Sub Highlight_Selected_Table_Row(ByVal Param1 As Object, Optional ByVal Param2 As Object = Nothing)
    On Error GoTo CleanExit

    Dim Target As Range
    Dim Table As ListObject

    If Param1 Is Nothing Then Exit Sub

    ' Resolve Target and Table flexibly from arguments (supports 1 or 2 params in any order)
    If TypeOf Param1 Is ListObject Then
        Set Table = Param1
        If Not Param2 Is Nothing Then
            If TypeOf Param2 Is Range Then Set Target = Param2
        End If
    ElseIf TypeOf Param1 Is Range Then
        Set Target = Param1
        If Not Param2 Is Nothing Then
            If TypeOf Param2 Is ListObject Then Set Table = Param2
        End If
    End If

    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 50 Then Exit Sub

    ' If Table not passed explicitly, resolve natively from Target
    If Table Is Nothing Then
        On Error Resume Next
        Set Table = Target.Cells(1, 1).ListObject
        On Error GoTo CleanExit
    End If

    Dim ws As Worksheet
    Set ws = Target.Worksheet

    ' If selection is outside table or table has no body, gracefully turn off highlight
    If Table Is Nothing Then
        DeactivateHighlight ws
        Exit Sub
    End If

    If Table.DataBodyRange Is Nothing Then Exit Sub

    If Intersect(Target, Table.DataBodyRange) Is Nothing Then
        DeactivateHighlight ws
        Exit Sub
    End If

    ' Ensure sheet-scoped ActiveTableRow name exists and points to selected row
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names("ActiveTableRow")
    On Error GoTo CleanExit

    If n Is Nothing Then
        ' Remove any conflicting workbook-level name
        On Error Resume Next
        ws.Parent.Names("ActiveTableRow").Delete
        On Error GoTo CleanExit
        Set n = ws.Names.Add(Name:="ActiveTableRow", RefersTo:="=" & Target.Row)
    Else
        If n.RefersTo <> "=" & Target.Row Then
            n.RefersTo = "=" & Target.Row
        End If
    End If

    ' Ensure conditional formatting rule exists, is unfragmented, and covers DataBodyRange
    EnsureTableHighlightRule Table

CleanExit:
    On Error GoTo 0
End Sub

' Backward-compatibility alias for existing callers with typo
Sub Highlight_Selected_Tabe_Row(ByVal Param1 As Object, Optional ByVal Param2 As Object = Nothing)
    Highlight_Selected_Table_Row Param1, Param2
End Sub

Private Sub DeactivateHighlight(ByVal ws As Worksheet)
    On Error Resume Next
    Dim n As Name
    Set n = ws.Names("ActiveTableRow")
    If Not n Is Nothing Then
        If n.RefersTo <> "=0" Then n.RefersTo = "=0"
    End If
    On Error GoTo 0
End Sub

Private Sub EnsureTableHighlightRule(ByVal Table As ListObject)
    On Error Resume Next
    Dim tblRange As Range
    Set tblRange = Table.DataBodyRange
    If tblRange Is Nothing Then Exit Sub

    Dim fc As FormatCondition
    Dim existingRule As FormatCondition
    Dim i As Long

    ' Iterate backwards to safely inspect and purge duplicate or fragmented rules
    For i = Table.DataBodyRange.FormatConditions.Count To 1 Step -1
        Set fc = Nothing
        Set fc = Table.DataBodyRange.FormatConditions(i)

        If Not fc Is Nothing Then
            ' Only xlExpression (type 2) has a Formula1 property
            If fc.Type = 2 Then
                Dim fText As String
                fText = ""
                fText = fc.Formula1

                If InStr(1, fText, "ActiveTableRow", vbTextCompare) > 0 Then
                    If existingRule Is Nothing Then
                        Set existingRule = fc
                    Else
                        ' Delete redundant duplicate rule
                        fc.Delete
                    End If
                End If
            End If
        End If
    Next i

    If existingRule Is Nothing Then
        ' Provision clean conditional formatting rule
        Set existingRule = tblRange.FormatConditions.Add(Type:=2, Formula1:="=ROW()=ActiveTableRow")
        With existingRule.Interior
            .Color = 10092543 ' Pale yellow
            .Pattern = 1      ' xlSolid
        End With
    Else
        ' Self-healing: adapt rule range if rows were added/deleted or range fragmented
        If existingRule.AppliesTo.Address <> tblRange.Address Then
            existingRule.ModifyAppliesToRange tblRange
        End If
    End If
    On Error GoTo 0
End Sub
