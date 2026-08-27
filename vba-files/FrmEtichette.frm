VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmEtichette 
   Caption         =   "Etichette"
   ClientHeight    =   2130
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   5652
   OleObjectBlob   =   "FrmEtichette.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmEtichette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CbAnnulla_Click()
    Unload FrmEtichetteColonna
End Sub

Private Sub CbOk_Click()
    Dim EtichetteColonna, EtichetteRiga As Range
    Dim NomeCella As String
    Dim CellaElaborata As Range
    Dim Riga, Colonna, Distanza As Integer

    Set EtichetteColonna = Range(RiferimentoColonne.Text)
    Set EtichetteRiga = Range(RiferimentoRighe.Text)
    Distanza = EtichetteRiga.Row - EtichetteColonna.Row - 1
    For Riga = 1 To EtichetteRiga.Cells.count
        For Colonna = 1 To EtichetteColonna.Cells.count
            If EtichetteColonna.Cells(Colonna).value <> "" Then
                If EtichetteRiga.Cells(Riga).value <> "" Then
                    NomeCella = stringasenzaspazi(EtichetteColonna.Cells(Colonna).value) & "_" & _
                                                                                         stringasenzaspazi(EtichetteRiga.Cells(Riga).value)
                    Set CellaElaborata = EtichetteColonna.Cells(Colonna).Offset(Riga + Distanza, 0)
                    ActiveWorkbook.names.Add Name:=NomeCella, RefersTo:=CellaElaborata
                    On Error GoTo 0
                End If
            End If
        Next Colonna
    Next Riga
    Unload FrmEtichette
End Sub


