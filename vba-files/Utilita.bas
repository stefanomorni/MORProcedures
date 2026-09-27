Attribute VB_Name = "Utilita"
'@Folder("MORProcedures.Modules")
'Declarations
Option Explicit

'Variable Declaration
Public EtichetteColonna, EtichetteRiga As Range
Public Annullato As Boolean
Dim Riga
Dim Nascoste, x, count, Lista, NoClienti As Integer
Dim Percorso, NomeFile, TuttiIClienti, BaseEstrazione, Criteri, Posizione, CasellaChiave As String
Dim TrovatoIn As Object

Sub AspettaXSecondi(ByVal Secondi As Integer)
    Dim Inizio As Date
    Inizio = Timer
    Do
    Loop Until (Timer - Inizio) >= Secondi
End Sub

Sub SeTrovatoBeepESeleziona(NomeBook As String, NomeSheet As String, Testo As String)
    
    Set TrovatoIn = Workbooks(NomeBook).Worksheets(NomeSheet).Cells.Find(Testo, , xlValues)
    If Not (TrovatoIn Is Nothing) Then
        Beep
        Workbooks(NomeBook).Worksheets(NomeSheet).Activate
        TrovatoIn.Offset(0, -1).Select
    End If
End Sub

Public Sub AssegnaNomiDaValoriRiga(ByVal Foglio As String, ByVal Suffisso As String, _
                                   ByVal NoRigaEtichette As Integer, _
                                   ByVal NoRigaOggetto As Integer)

    'Crea i nomi delle caselle della riga oggetto = Suffisso&[Valorecelle] della riga etichetta


    Dim Contatore As Integer
    Dim Etichetta As String
    Dim MetodoCalcolo

    MetodoCalcolo = Application.Calculation
    Application.Calculation = xlCalculationManual
    For Contatore = 1 To Worksheets(Foglio).Rows(NoRigaOggetto).Cells.count
        Etichetta = Worksheets(Foglio).Cells(NoRigaEtichette, Contatore).Text
        On Error Resume Next
        ActiveWorkbook.Names.Add Name:=stringasenzaspazi(Suffisso) & Etichetta, _
                                                                   RefersTo:=Worksheets(Foglio).Cells(NoRigaOggetto, Contatore), Visible:=True
    Next Contatore
    Application.Calculation = MetodoCalcolo


End Sub

Public Sub AssegnaNomiACellaTabellaDaEtichette()
    FrmEtichette.Show
End Sub

Public Sub Estrai_Nomi_Workbook(Libro As Workbook)
    Dim Nomi As Collection
    Dim Indirizzi As Collection
    Nomi = Libro.Names.Item.Name
    Indirizzi = Libro.Names.Item.RefersToRange
End Sub
