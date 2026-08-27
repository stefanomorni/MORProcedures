Attribute VB_Name = "Grafici"
'Direttive

Option Explicit                                  'rende obbligatoria la dichiarazione delle variabili
Option Compare Text                              'paragona su base testo piuttosto che binaria
Option Base 1                                    'stabilisce 1 come punto di partenza nei For...Next

Public Sub UpdateWbChartsSourceData(ByVal wb As Object)
Dim Cs As Object
Dim ws As Object
On Error GoTo Messaggio
For Each Cs In wb.Charts
    UpdateCsChartsSourceData Cs
Next Cs
For Each ws In wb.Worksheets
    UpdateWsChartsSourceData ws
Next ws
On Error GoTo 0
Exit Sub
Messaggio:
MsgBox "Non ho potuto aggiornare il chart " & Cs.Name & " ed ho interrotto l'aggiornamento"
End Sub

Public Sub UpdateWsChartsSourceData(ByVal ws As Object)
Dim Cht As Object
On Error GoTo Messaggio
For Each Cht In ws.ChartObjects
    Cht.Chart.SetSourceData Source:=rgSourceDataregion(Cht.Chart)
Next Cht
On Error GoTo 0
Exit Sub
Messaggio:
MsgBox "Non ho potuto aggiornare il chart " & Cht.Name & " ed ho interrotto l'aggiornamento"
End Sub
Public Sub UpdateCsChartsSourceData(ByVal Cs As Object)
Dim Cht As Object
On Error GoTo Messaggio
Cs.SetSourceData Source:=rgSourceDataregion(Cs)
On Error GoTo 0
Exit Sub
Messaggio:
MsgBox "Non ho potuto aggiornare il chart " & Cht.Name & " ed ho interrotto l'aggiornamento"
End Sub

Public Sub UpdateWbChartsLabelAndAxes(ByVal wb As Object, ShowName As Boolean, Separator As String, Optional LabelFormat = "", Optional FontSize As Double = 0)
Dim Cht As Object
Dim ws As Object

On Error GoTo Messaggio
For Each Cht In wb.Charts
    UpdateCsChartLabelAndAxes Cht, ShowName, Separator, LabelFormat, FontSize
Next Cht
For Each ws In wb.Worksheets
    For Each Cht In ws.ChartObjects.Charts
        Cht.SetSourceData Source:=rgSourceDataregion(Cht.Chart)
    Next Cht
Next ws
On Error GoTo 0
Exit Sub
Messaggio:
MsgBox "Non ho potuto aggiornare le etichette del chart " & Cht.Name & " ed ho interrotto l'aggiornamento"
End Sub

Public Sub UpdateCsChartLabelAndAxes(ByVal Cs As Object, ShowName As Boolean, Optional Separator As String = " ", Optional LabelFormat = "", Optional FontSize As Double = 0, Optional PctVerticalMargin As Double = 0, Optional AxesMainUnit As Double = 0.05, Optional MiniorUnitPerMainUnit As Double = 5)
Dim Serie As Object
Dim Min
Dim Max

Min = MinCht(Cs)
Max = MaxCht(Cs)
'On Error GoTo Messaggio
If Min > 0 Then
    Cs.Axes(xlValue).MinimumScale = Round((Min * (100 - PctVerticalMargin) / 100) / AxesMainUnit) * AxesMainUnit
Else
    Cs.Axes(xlValue).MinimumScale = -1 * (Round(Abs((Min) * (100 + PctVerticalMargin) / 100) / AxesMainUnit) * AxesMainUnit)
End If
If Max > 0 Then
    Cs.Axes(xlValue).MaximumScale = Round((Max * (100 + PctVerticalMargin) / 100) / AxesMainUnit) * AxesMainUnit
Else
    Cs.Axes(xlValue).MaximumScale = -1 * ((Round(Abs(Max) * (100 + PctVerticalMargin) / 100) / AxesMainUnit) * AxesMainUnit)
End If
Cs.Axes(xlValue).MajorUnit = AxesMainUnit
Cs.Axes(xlValue).MinorUnit = AxesMainUnit / MiniorUnitPerMainUnit
LabelLastPoint Cs, ShowName, Separator, LabelFormat, FontSize
On Error GoTo 0
Exit Sub
Messaggio:
MsgBox "Non ho potuto aggiornare le etichette del chart " & Cs.Name & " ed ho interrotto l'aggiornamento"
End Sub

Public Sub Resize_Canvas(Cht As Object, Width_in_cm As Long, Height_in_cm As Long, Line_weight_in_points)
Dim Max_DataLabel_Width As Long
Dim Serie As Object
Max_DataLabel_Width = 0
For Each Serie In Cht.FullSeriesCollection
    With Serie.Points(Serie.Points.count)
        If .HasDataLabel Then
            On Error Resume Next
            Dim lblWidth As Long
            lblWidth = .DataLabel.Width
            If Err.Number = 0 Then
                If lblWidth > Max_DataLabel_Width Then
                    Max_DataLabel_Width = lblWidth
                End If
            End If
            On Error GoTo 0
        End If
    End With
Next Serie
With Cht.PlotArea
    .Width = Application.CentimetersToPoints(Width_in_cm) - Max_DataLabel_Width
    .Height = Application.CentimetersToPoints(Height_in_cm) - Cht.Axes(xlCategory).Height
End With
For Each Serie In Cht.FullSeriesCollection
    Serie.Format.line.Weight = Line_weight_in_points
Next Serie
End Sub

Public Sub LabelLastPoint(ByVal Cht As Object, ByVal ShowName As Boolean, ByVal Separator As String, Optional ByVal LabelFormat As String = "", Optional ByVal FontSize As Double = 0)
    Dim Serie As Object
    Dim last_available_point As Long
    Dim Values As Variant

    Dim plotTop As Double, plotBottom As Double
    Dim lbls() As DataLabel
    Dim sCount As Long, sIdx As Long

    ' Plot area bounds for clamping
    plotTop = Cht.PlotArea.InsideTop
    plotBottom = plotTop + Cht.PlotArea.InsideHeight

    ' Clear existing labels
    Cht.SetElement (msoElementDataLabelNone)

    ' Prepare array to store each series' last-point label
    sCount = Cht.SeriesCollection.count
    If sCount = 0 Then Exit Sub
    ReDim lbls(1 To sCount)

    ' Apply labels to last available point in each series and store the exact label reference
    For sIdx = 1 To sCount
        Set Serie = Cht.SeriesCollection(sIdx)
        last_available_point = Serie.Points.count
        Values = Serie.Values

        ' Find last non-empty point (with bounds guard)
        Do While last_available_point > 0
            If Not IsError(Values(last_available_point)) And Not IsEmpty(Values(last_available_point)) And IsNumeric(Values(last_available_point)) Then
                Exit Do
            End If
            last_available_point = last_available_point - 1
        Loop

        If last_available_point > 0 Then
            With Serie.Points(last_available_point)
                .ApplyDataLabels
                With .DataLabel
                    .ShowValue = True
                    .ShowSeriesName = ShowName
                    .Separator = Separator

                    ' Choose a movable position (often required for reliable Top/Left changes)
                    ' You can change to xlLabelPositionAbove/Below depending on chart type preference
                    .Position = xlLabelPositionRight

                    If FontSize > 0 Then
                        .Format.TextFrame2.TextRange.Font.Size = FontSize
                        .Format.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
                    End If

                    .AutoText = True

                    If Len(LabelFormat) = 0 Then
                        .NumberFormat = Cht.Axes(xlValue).TickLabels.NumberFormat
                    Else
                        .NumberFormat = LabelFormat
                    End If
                End With

                ' Store the exact last-point label for overlap adjustment
                Set lbls(sIdx) = .DataLabel
            End With
        Else
            Set lbls(sIdx) = Nothing
        End If
    Next sIdx

    ' Now remove vertical overlaps among last-point labels
    AdjustLastPointLabelOverlaps Cht, lbls, plotTop, plotBottom
End Sub

' Move labels to eliminate vertical overlap, clamping within plot area bounds.
Private Sub AdjustLastPointLabelOverlaps(ByVal Cht As Object, ByRef lbls() As DataLabel, ByVal plotTop As Double, ByVal plotBottom As Double)
    Dim i As Long, j As Long
    Dim n As Long
    Dim moved As Boolean
    Dim overlap As Double
    Dim newTop As Double
    Dim safetyPass As Long
    Const MaxPasses As Long = 10       ' cap the number of passes
    Const Padding As Double = 2        ' small gap to avoid touching

    n = UBound(lbls)
    If n = 0 Then Exit Sub

    ' Multi-pass simple resolver: compare each pair and nudge down the lower label
    For safetyPass = 1 To MaxPasses
        moved = False

        For i = 1 To n
            If Not IsLabelPositionable(lbls(i)) Then GoTo NextI

            For j = i + 1 To n
                If Not IsLabelPositionable(lbls(j)) Then GoTo NextJ

                ' Check vertical overlap (bounding boxes intersect vertically)
                If BoxesOverlapVertically(lbls(i), lbls(j)) And BoxesOverlapHorizontally(lbls(i), lbls(j)) Then
                    ' Decide which label to move: push the lower one further down
                    If lbls(i).Top >= lbls(j).Top Then
                        overlap = (lbls(j).Top + lbls(j).Height) - lbls(i).Top
                        If overlap > 0 Then
                            newTop = lbls(i).Top + overlap + Padding
                            ' Clamp
                            If newTop < plotTop Then newTop = plotTop
                            If newTop + lbls(i).Height > plotBottom Then newTop = plotBottom - lbls(i).Height
                            If newTop <> lbls(i).Top Then
                                lbls(i).Top = newTop
                                moved = True
                            End If
                        End If
                    Else
                        overlap = (lbls(i).Top + lbls(i).Height) - lbls(j).Top
                        If overlap > 0 Then
                            newTop = lbls(j).Top + overlap + Padding
                            ' Clamp
                            If newTop < plotTop Then newTop = plotTop
                            If newTop + lbls(j).Height > plotBottom Then newTop = plotBottom - lbls(j).Height
                            If newTop <> lbls(j).Top Then
                                lbls(j).Top = newTop
                                moved = True
                            End If
                        End If
                    End If
                End If
NextJ:
            Next j

NextI:
        Next i

        If Not moved Then Exit For
    Next safetyPass
End Sub

' Helper: Check if a DataLabel can be positioned (exists and exposes Top/Height safely)
Private Function IsLabelPositionable(ByVal lbl As DataLabel) As Boolean
    On Error GoTo notPos
    If lbl Is Nothing Then
        IsLabelPositionable = False
        Exit Function
    End If
    Dim t As Double, h As Double
    t = CDbl(lbl.Top)
    h = CDbl(lbl.Height)
    IsLabelPositionable = (h > 0)
    Exit Function
notPos:
    IsLabelPositionable = False
End Function

' Helper: axis-aligned bounding box intersection (vertical)
Private Function BoxesOverlapVertically(ByVal a As DataLabel, ByVal b As DataLabel) As Boolean
    On Error Resume Next
    BoxesOverlapVertically = (a.Top < (b.Top + b.Height)) And ((a.Top + a.Height) > b.Top)
End Function

' Helper: axis-aligned bounding box intersection (horizontal)
Private Function BoxesOverlapHorizontally(ByVal a As DataLabel, ByVal b As DataLabel) As Boolean
    On Error Resume Next
    BoxesOverlapHorizontally = (a.Left < (b.Left + b.Width)) And ((a.Left + a.Width) > b.Left)
End Function

Sub SelezionaGraficiWorkbook(Optional Azione As String)

    Dim Pagina As Object
        
    For Each Pagina In ActiveWorkbook.Sheets
        Pagina.Shapes.SelectAll
        If Azione = "Separa" Then
            On Error Resume Next                 'se non ci sonomgrafici nel foglio
            Selection.Ungroup.Select
        ElseIf Azione = "RiRagruppa" Then
            On Error Resume Next                 'se non ci sono grafici nel foglio
            Selection.Regroup.Select
        End If
    Next Pagina
 
End Sub

Public Sub SeparaGraficiDelWorkbook()
    SelezionaGraficiWorkbook "Separa"
End Sub

Public Sub RiRaggruppaGraficiDelWorkbook()
    SelezionaGraficiWorkbook "Riragruppa"
End Sub

Public Sub CercaNelleSerie()

    Dim Contatore, Contatore1, Contatore2 As Integer
    Dim StringaRicercata, Prova As String
    Dim NumeroFogli, NumeroGrafici, NumeroSerie As Long
    Dim PosizioneNelRiferimento As Integer
    Dim MsgTitolo, MsgTesto As String
    Dim MsgPulsanti As Integer
    Dim Risposta As Integer
    Dim Serie
    Dim RiferimentiSerie, NomeSerie, TitoloGrafico As String
    
    SeparaGraficiDelWorkbook
        
    
    MsgTitolo = "Ritrovamento Stringa"
    MsgPulsanti = vbAbortRetryIgnore + vbInformation + vbDefaultButton1
   
    
    StringaRicercata = InputBox("Inserisci la stringa da ricercare nei dati di origine", "Ricerca stinga nei grafici")
    With ActiveWorkbook
        NumeroFogli = Worksheets.count
        For Contatore = 1 To NumeroFogli
            With Worksheets(Contatore)
                NumeroGrafici = .ChartObjects.count
                For Contatore1 = 1 To NumeroGrafici
                    NumeroSerie = .ChartObjects(Contatore1).Chart.SeriesCollection.count
                    For Contatore2 = 1 To NumeroSerie
                        On Error Resume Next
                        RiferimentiSerie = Worksheets(Contatore).ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Formula
                        NomeSerie = .ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Name
                        TitoloGrafico = .ChartObjects(Contatore1).Chart.ChartTitle.Caption
                        PosizioneNelRiferimento = InStr(RiferimentiSerie, StringaRicercata)
                        If PosizioneNelRiferimento <> 0 Then
                            MsgTesto = "La stringa ricercata è stata trovata nei riferimenti della serie " & NomeSerie & " del Grafico  " & TitoloGrafico & " del foglio " & Worksheets(Contatore).Name
                            Risposta = MsgBox(MsgTesto, MsgPulsanti, MsgTitolo)
                            If Risposta = 3 Then
                                ActiveCell.Range("A1:C1").Select
                                .ChartObjects(Contatore1).Activate
                                ActiveChart.ChartArea.Select
                                Exit Sub
                            ElseIf Risposta = 5 Then
                                Risposta = MsgBox("La ricerca è stata annulata", vbInformation + vbOKOnly, MsgTitolo)
                                Exit Sub
                            End If
                        End If
                        On Error GoTo 0
                    Next Contatore2
                Next Contatore1
            End With                             'worksheets(contatore)
        Next Contatore
    End With                                     'Active workbook
    Risposta = MsgBox("La stringa ricercata non è stata trovata", vbInformation + vbOKOnly, MsgTitolo)

    'RiRaggruppaGraficiDelWorkbook

End Sub

Public Sub SostituisciNelleSerie()
    Dim Contatore, Contatore1, Contatore2, NoSostituzioni As Integer
    Dim StringaRicercata, StringaSostitutiva, Prova As String
    Dim NumeroFogli, NumeroGrafici, NumeroSerie As Long
    Dim PosizioneNelRiferimento As Integer
    Dim MsgTitolo, MsgTesto As String
    Dim MsgPulsanti As Integer
    Dim Risposta As Integer
    Dim Serie
    Dim RiferimentiSerie, NomeSerie, TitoloGrafico As String
    
    SeparaGraficiDelWorkbook
        
    
    MsgTitolo = "Stringa da sostituire"
    MsgPulsanti = vbAbortRetryIgnore + vbInformation + vbDefaultButton1
    StringaRicercata = InputBox("Inserisci la stringa da sostituire nei dati di origine", "Ricerca stinga nei grafici")
    MsgTitolo = "Stringa Sostitutiva"
    MsgPulsanti = vbRetryCancel + vbInformation + vbDefaultButton1
    StringaSostitutiva = InputBox("Inserisci la stringa sostitutiva nei dati di origine", "Ricerca stinga nei grafici")
    NoSostituzioni = 0
    With ActiveWorkbook
        NumeroFogli = Worksheets.count
        For Contatore = 1 To NumeroFogli
            With Worksheets(Contatore)
                NumeroGrafici = .ChartObjects.count
                For Contatore1 = 1 To NumeroGrafici
                    NumeroSerie = .ChartObjects(Contatore1).Chart.SeriesCollection.count
                    For Contatore2 = 1 To NumeroSerie
                        On Error Resume Next
                        RiferimentiSerie = Worksheets(Contatore).ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Formula
                        NomeSerie = .ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Name
                        TitoloGrafico = .ChartObjects(Contatore1).Chart.ChartTitle.Caption
                        PosizioneNelRiferimento = InStr(RiferimentiSerie, StringaRicercata)
                        If PosizioneNelRiferimento <> 0 Then
                            MsgTesto = "La stringa ricercata è stata trovata nei riferimenti della serie " & NomeSerie & " del Grafico  " & TitoloGrafico & " del foglio " & Worksheets(Contatore).Name & ". Essa verrà sostituita con " & StringaSostitutiva
                            Risposta = MsgBox(MsgTesto, MsgPulsanti, MsgTitolo)
                            If Risposta = 5 Then
                                Risposta = MsgBox("La ricerca è stata annulata; sono state effettuate " & NoSostituzioni & " sostituzioni.", vbInformation + vbOKOnly, MsgTitolo)
                                Exit Sub
                            End If
                            Worksheets(Contatore).ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Formula = Replace(Worksheets(Contatore).ChartObjects(Contatore1).Chart.SeriesCollection(Contatore2).Formula, StringaRicercata, StringaSostitutiva)
                            NoSostituzioni = NoSostituzioni + 1
                        End If
                        On Error GoTo 0
                    Next Contatore2
                Next Contatore1
            End With                             'worksheets(contatore)
        Next Contatore
    End With                                     'Active workbook
    If PosizioneNelRiferimento <> 0 Then
        Risposta = MsgBox("L'oprazione è terminata. Sono state eseguite " & NoSostituzioni & " sostituzioni.", vbInformation)
    Else
        Risposta = MsgBox("La stringa ricercata non è stata trovata", vbInformation + vbOKOnly, MsgTitolo)
    End If
    'RiRaggruppaGraficiDelWorkbook

End Sub

Public Sub Sort_Series_On_Last_Value(ByVal Cht As Object, Optional ByVal Ascending As Boolean = False)
    Dim n As Long, i As Long
    Dim vals As Variant
    Dim lastIdx As Long
    Dim lastVal As Double
    Dim seriesName As String
    Dim pairs() As Variant  ' pairs(i, 0) = value (Double), pairs(i, 1) = name (String)

    With Cht
        n = .SeriesCollection.count
        If n = 0 Then Exit Sub

        ReDim pairs(1 To n, 0 To 1)

        For i = 1 To n
            seriesName = .SeriesCollection(i).Name
            vals = .SeriesCollection(i).Values

            ' Find the last non-empty numeric value by scanning backward
            lastIdx = GetLastNumericIndex(vals)
            If lastIdx > 0 Then
                lastVal = CDbl(vals(lastIdx))
            Else
                ' If no numeric value found, push it to the bottom by using -Inf surrogate
                lastVal = -1E+308
            End If

            pairs(i, 0) = lastVal
            pairs(i, 1) = seriesName
        Next i

        ' Sort the pairs by value
        SortPairsByValue pairs, Ascending

        ' Apply plot order according to sorted names
        For i = 1 To n
            .SeriesCollection(CStr(pairs(i, 1))).PlotOrder = i
        Next i
    End With
End Sub

' Returns the index of the last element in vals that is numeric and not Empty/Error; 0 if none.
Private Function GetLastNumericIndex(ByVal vals As Variant) As Long
    Dim lb As Long, ub As Long, j As Long
    On Error Resume Next

    lb = LBound(vals)
    ub = UBound(vals)

    For j = ub To lb Step -1
        If IsError(vals(j)) Then GoTo ContinueLoop
        If IsEmpty(vals(j)) Then GoTo ContinueLoop
        If IsNumeric(vals(j)) Then
            GetLastNumericIndex = j
            Exit Function
        End If
ContinueLoop:
    Next j

    GetLastNumericIndex = 0
End Function

' In-place sort of pairs(1..n, 0..1) by column 0 (Double value)
Private Sub SortPairsByValue(ByRef pairs() As Variant, ByVal Ascending As Boolean)
    Dim n As Long, i As Long, j As Long
    Dim v1 As Double, v2 As Double
    Dim tmpVal As Double, tmpName As String

    On Error Resume Next
    n = UBound(pairs, 1)

    ' Stable-ish bubble sort for simplicity; n is small (series count)
    For i = 1 To n - 1
        For j = 1 To n - i
            v1 = CDbl(pairs(j, 0))
            v2 = CDbl(pairs(j + 1, 0))
            If Ascending Then
                If v1 > v2 Then
                    tmpVal = pairs(j, 0)
                    tmpName = CStr(pairs(j, 1))
                    pairs(j, 0) = pairs(j + 1, 0)
                    pairs(j, 1) = pairs(j + 1, 1)
                    pairs(j + 1, 0) = tmpVal
                    pairs(j + 1, 1) = tmpName
                End If
            Else
                If v1 < v2 Then
                    tmpVal = pairs(j, 0)
                    tmpName = CStr(pairs(j, 1))
                    pairs(j, 0) = pairs(j + 1, 0)
                    pairs(j, 1) = pairs(j + 1, 1)
                    pairs(j + 1, 0) = tmpVal
                    pairs(j + 1, 1) = tmpName
                End If
            End If
        Next j
    Next i
End Sub


