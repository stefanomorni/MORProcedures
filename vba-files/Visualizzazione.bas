Attribute VB_Name = "Visualizzazione"
'@Folder("MORProcedures.Modules")
Option Explicit

Sub Fissa_AreaVisualizzazione_Su_Selezione_Corrente()
    '
    ' Fissa_AreaVisualizzazione_Su_Selezione_Corrente Macro
    ' Attribuisce il nome [nomefolgio]_AreaVisualizzazione alle selezioni correnti dei diversi folgli. Tale Nome serve da base per la macro Zoom_AreaVisualizzazione
    '

    'Declarations

    Dim Foglio As Object

    For Each Foglio In ActiveWorkbook.Worksheets
        Foglio.Activate
        ActiveWorkbook.Names.Add Name:=ActiveSheet.Name & "_AreaVisualizzazione", _
                                 RefersTo:=Selection
    Next Foglio

End Sub

Sub Zooma_Su_AreaVisualizzazione(ByVal sh As Object)

    ' Aggiusta automaticamente lo Zoom in modo da visualizzare l'area
    ' [NomeFoglio]_AreaVisualizzazione sul foglio fornito come argomento. Tali nomi possono essere creati per tutti
    ' i fogli del libro corrente attraverso la MacroFissa_AreaVisualizzazione_Su_Selezione_Corrente()

    Dim SelezionePrecedente As String

    SelezionePrecedente = ActiveCell.Address

    Application.GoTo Reference:=ActiveSheet.Name & "_AreaVisualizzazione"
    ActiveWindow.Zoom = True
    Range(SelezionePrecedente).Activate

End Sub

Sub Incrementa_Spazio_Righe_Tabelle(ByRef Foglio As Object, Punti As Integer)
    Dim Riga As Range
    Dim TblTabella, pvtTabella As Object
    Application.ScreenUpdating = False
    For Each TblTabella In Foglio.ListObjects
        If TblTabella.DataBodyRange Is Nothing Then GoTo fine_tbl
        For Each Riga In TblTabella.DataBodyRange.EntireRow
            If Riga.Hidden = False Then
                Riga.AutoFit
                Riga.RowHeight = Riga.RowHeight + Punti
            End If
        Next Riga
fine_tbl:
    Next TblTabella
    For Each pvtTabella In Foglio.PivotTables
        If pvtTabella.DataBodyRange Is Nothing Then GoTo fine_pvt
        For Each Riga In pvtTabella.DataBodyRange.EntireRow
            If Riga.Hidden = False Then
                Riga.AutoFit
                Riga.RowHeight = Riga.RowHeight + Punti
            End If
        Next Riga
fine_pvt:
    Next pvtTabella
    Application.ScreenUpdating = True
End Sub

Sub Incrementa_Spazio_Righe_Area(area As Object, Punti As Integer)
    Dim Riga As Range
    Application.ScreenUpdating = False
    For Each Riga In area.Rows.EntireRow
        If Riga.Hidden = False Then
            Riga.AutoFit
            Riga.RowHeight = Riga.RowHeight + Punti
        End If
    Next Riga
    Application.ScreenUpdating = True
End Sub
