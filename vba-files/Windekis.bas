Attribute VB_Name = "Windekis"
'@Folder("MORProcedures.Modules")
Public Sub Importa_Linee_Windekis()
    Dim c As Integer
    Dim NoLinee As Integer
    Dim CellaAttiva As Range

    ' Richiedi no di linee
    'NoLinee = InputBox("Inserisci il numero di linee da importare (in alto a destra in Windekis)", "Numero Linee importazione")
    'Accoda i data alla tabella corrente
    Set CellaAttiva = ActiveWorkbook.ActiveSheet.Rows(ActiveSheet.UsedRange.Rows.count + 1).Cells(1, 1)

    'Copia le righe 1 a 1

    VBA.Interaction.AppActivate "Windekis"
    Do
        VBA.Interaction.SendKeys "^c", True
        VBA.Interaction.SendKeys "{DOWN}", True
        DoEvents
        CellaAttiva.Select
        ActiveSheet.Paste
        Set CellaAttiva = CellaAttiva.Offset(1, 0)
    Loop
End Sub
