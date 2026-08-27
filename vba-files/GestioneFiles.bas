Attribute VB_Name = "GestioneFiles"
Public Sub SeWorkbookNonApertoApri(NomeCompletoFile As String)
    Dim FileAperto As Boolean
    FileAperto = False                           ' Inizializza la variabile.
    For Each Workbook In Workbooks               ' Esegue un'iterazione in ogni elemento.
        If Workbook.FullName = NomeCompletoFile Then
            Exit Sub
        End If
    Next
    Workbooks.Open Filename:=NomeCompletoFile
End Sub

Public Sub EseguiSuTuttiIWokbook(FullPathCartella As String, NomeProcedura As String)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If Right(FullPathCartella, 1) <> Application.PathSeparator Then
        FullPathCartella = FullPathCartella & Application.PathSeparator
    End If
    Fname = Dir(FullPathCartella & "*.xl*")      'loop through the files
    Do While Len(Fname)
        SeWorkbookNonApertoApri FullPathCartella & Fname
        With ActiveWorkbook                      '
            Application.Run NomeProcedura
            Workbooks(Fname).Close SaveChanges:=True
        End With                                 ' go to the next file in the folder
        Fname = Dir
    Loop
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Public Sub ExportRange(area As Object, Format As Long, Optional Fullpath As String, Optional Filename As String)

    Dim wb As Workbook, wbNew As Workbook
    Dim ws As Worksheet, wsNew As Worksheet
    Dim wbNewName, wbSavePath As String



    MessaggioInBarraDiStato "Sto esportando il range. " & "Un attimo di pazienza p.f.", 5
    Application.ScreenUpdating = False
    Set wb = area.Parent.Parent
    Set wbNew = Workbooks.Add

    If Fullpath <> "" Then wbSavePath = Fullpath Else wbSavePath = wb.Path
    If Filename <> "" Then wbNewName = Filename Else Filename = area.Name
    Set wsNew = wbNew.Sheets(1)
    area.Copy
    wsNew.Range("A1").PasteSpecial Paste:=xlPasteAll
    SaveWorkbook wbNew, wbSavePath, Filename, Format, False, True
    If Not (Application.VBE.MainWindow.Visible) Then
        Application.VBE.MainWindow.Visible = True
        Application.VBE.MainWindow.Visible = False
    End If
    wbNew.Close False
    'Set wbNew = Nothing
    Application.ScreenUpdating = True
    ResetBarraDiStato

End Sub

Public Sub SaveWorkbook(File As Object, Fullpath As String, Filename As String, FileFormat, Createbackup As Boolean, Overwrite As Boolean)
    Dim InitialState As Boolean
    InitialState = Application.DisplayAlerts
    MessaggioInBarraDiStato "Stò salvando il file " & Fullpath & "\" & Filename & ". Un attimo di pazienza p.f.", 5
    If Overwrite = True Then Application.DisplayAlerts = False Else Application.DisplayAlerts = True
    On Error GoTo Messaggio
    File.SaveAs Filename:=Fullpath & "\" & Filename, FileFormat:=FileFormat, Createbackup:=Createbackup
    Application.DisplayAlerts = InitialState
    ResetBarraDiStato
    Exit Sub
Messaggio:
    MsgBox "Portroppo non ho potuto salvare il file " & Fullpath & "\" & Filename & ". P.f. Verifica il nome del file e delal directory o gli accessi di scrittura.", vbCritical
    Resume Next
End Sub








