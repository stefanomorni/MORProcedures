Attribute VB_Name = "Editing"
Option Explicit

Public Sub Array2Range(ByVal objArray, FirstCell, Optional Transpose As Boolean = False, _
                        Optional RangeName As String = "", Optional FormatAsTable As Boolean = True)
Dim r, c As Integer
Dim Destination As Range

'Determine rows and columns of array
r = UBound(objArray, 1) - LBound(objArray, 1)
c = UBound(objArray, 2) - LBound(objArray, 2)
'Check or transform Firstcell to range
If Not IsObject(FirstCell) Then Set FirstCell = Range(FirstCell)
If Transpose Then
    Set Destination = FirstCell.Cells(1, 1).Resize(c, r)
    Destination.value = Application.Transpose(objArray)
Else
    Set Destination = FirstCell.Cells(1, 1).Resize(r, c)
    Destination.value = objArray
End If

If FormatAsTable Then
    ActiveSheet.ListObjects.Add(xlSrcRange, Destination, , xlYes).Name = "ArrayTable"
    If RangeName <> "" Then
        ActiveSheet.ListObjects("ArrayTable").Name = RangeName
    End If
ElseIf RangeName <> "" Then
    ActiveWorkbook.names.Add RangeName, Destination
End If
End Sub

Public Sub Fix_Range_Values_PasteDown_Formulas(area As String)
'Application.ScreenUpdating = False
Range(area).Copy (Selection.Offset(1, 0))
With Range(area)
    .Copy
    .PasteSpecial Paste:=xlPasteValues
End With
Application.CutCopyMode = False
'Application.ScreenUpdating = True
End Sub

Public Sub Cancella_Tutti_I_Nomi_Del_Worbook()

    Dim NoNomi, Contatore As Integer
    NoNomi = ActiveWorkbook.names.count
    For Contatore = NoNomi To 1 Step -1
        ActiveWorkbook.names(Contatore).Delete
    Next Contatore

End Sub

Public Sub CentraSuSelezione()
    With Selection

        If .HorizontalAlignment = xlCenterAcrossSelection Then
            .HorizontalAlignment = xlGeneral
        Else
            .HorizontalAlignment = xlCenterAcrossSelection
            .VerticalAlignment = xlCenter
        End If
    End With
End Sub

Public Sub Define_UsedRange()
    '
    ' Attibuisce il Nome "UsedRange" all UsedRange del Foglio

    Dim Foglio As Object

    With ActiveWorkbook
        For Each Foglio In .Worksheets
            ActiveWorkbook.names.Add Name:="UsedRange", RefersToR1C1:= _
                                     Foglio.UsedRange
        Next Foglio
    End With
End Sub

Public Sub Add_Column(ColNo As Integer, NewHeading As String)
    Dim Foglio As Object

    With ActiveWorkbook.ActiveSheet
        .Columns(ColNo).Select
        Selection.Insert Shift:=xlToRight
        .Columns(ColNo).Cells(1, 1).value = NewHeading
    End With
End Sub

Public Sub Aggiorna_Hyperlinks(Oggetto As Range)
    Dim Collegamenti, clCollegamento
    Dim strFormula As String
    'Ambiente
    Application.ScreenUpdating = False
    On Error GoTo fine

    Set Oggetto = ActiveCell.Parent
    Set Collegamenti = Oggetto.FindAll("http", Oggetto, xlFormulas)
    If Not Collegamenti Is Nothing Then
        For Each clCollegamento In Collegamenti.Cells
            strFormula = clCollegamento.Formula
            clCollegamento.ClearContents
            DoEvents
            If TypeOf Oggetto Is Worksheet Then
                Oggetto.Hyperlinks.Add Anchor:=clCollegamento, Address:=strFormula, TextToDisplay:=strFormula
            ElseIf TypeOf Oggetto Is Range Then
                Oggetto.parentHyperlinks.Add Anchor:=clCollegamento, Address:=strFormula, TextToDisplay:=strFormula
            End If
        Next clCollegamento
    End If

    Set Collegamenti = Oggetto.FindAll("HYPERLINK", Oggetto, xlFormulas)
    If Not Collegamenti Is Nothing Then
        For Each clCollegamento In Collegamenti.Cells
            strFormula = clCollegamento.Formula
            clCollegamento.ClearContents
            DoEvents
            clCollegamento.Formula = strFormula
        Next clCollegamento
    End If

fine:
    'Ambiente
    Application.ScreenUpdating = True

End Sub

Public Sub NumberAsText2Text()
    Dim Cella, area  As Range
    Dim Valore As String
    Dim Conta, Risposta, NoCelle  As Integer
    Dim Calcolo

    NoCelle = Intersect(Selection, ActiveSheet.UsedRange).Cells.count
    Risposta = MsgBox("Stò per controllare " & NoCelle & " celle. Se necessario riformaterò i numeri formatati come testo in testo. Vuoi continuare?", vbYesNo)
    Conta = 0
    Calcolo = Application.Calculation
    If Risposta = 6 Then
        Application.Calculation = xlCalculationManual
        For Each Cella In Intersect(Selection, ActiveSheet.UsedRange).Cells
            If Cella.NumberFormat = "@" And IsNumeric(Cella) And Cella <> "" Then
                Valore = CStr(Cella.value)
                Cella.value = Valore
                Conta = Conta + 1
            End If
        Next Cella
        Application.Calculation = Calcolo
        MsgBox ("Ho convertito in testo " & Conta & " numeri formatati come testo. Grazie per la pazienza.")
    End If
End Sub


