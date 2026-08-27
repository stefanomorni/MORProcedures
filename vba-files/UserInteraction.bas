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
    Dim RelevantRow As Range

    On Error GoTo fine
    Application.ScreenUpdating = False
    If Not IsError(Intersect(Selezione, Table.DataBodyRange)) Then          ' in case is a selection in table
        'Restore previous pattern/ color formating
        With Table.DataBodyRange.Interior ' XXX
            .Pattern = xlNone
        End With
        'highlight with pale yellow

        With Application.Intersect(Selezione.EntireRow, Table.DataBodyRange).Interior
            .Pattern = xlSolid
            '.PatternColorIndex = xlAutomatic
            .Color = 10092543
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
fine:
On Error GoTo 0
Application.ScreenUpdating = True
End Sub
