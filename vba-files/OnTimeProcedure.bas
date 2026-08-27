Attribute VB_Name = "OnTimeProcedure"
'@Folder("MORProcedures.Modules")
Option Explicit

Public Sub EseguiOgniXSecondi(Optional Libro As Workbook, Optional procedure As String, Optional SecondsFrequency As Single)
    '
    ' Runs the"procedure" every "Frequency" minutes from first call on the "CurrentBook" as long as it is still active; the "active check"
    ' every "frequency" minutes remains in action until the "Currentbook" is closed

    'Declarations

    Static CurrentBook As Workbook
    Static Procedura As String
    Static SecondiFrequenza As Single

    Dim CurrentProcedure As String
    Dim found As Boolean
    Dim UpdateTime As Date
    Dim BlockTime As Date

    'Initialisations
 
    Let CurrentProcedure = "EseguiOgniXSecondi"



    'Procedure

    If procedure <> "" Then
        Let Procedura = procedure
        Set CurrentBook = Libro
    End If
    On Error GoTo Filechiuso
    CurrentBook.Activate
    Application.Run ("'" & Libro.Name & "'!" & procedure)
    'Tests values
    If IsMissing(SecondsFrequency) Or SecondsFrequency = 0 Then
        If SecondiFrequenza < 1 Then SecondiFrequenza = 1
    Else
        SecondiFrequenza = SecondsFrequency
    End If
    UpdateTime = Now() + TimeValue("00:00:" & SecondiFrequenza)
    Application.OnTime EarliestTime:=UpdateTime, procedure:=CurrentProcedure

Filechiuso:
End Sub

Public Sub PausaMacroXSecondi(ByVal Minuti As Integer, ByVal Secondi As Integer)

    Dim NuovaOra, NuovoMinuti, NuovoSecondi, Pausa

    NuovaOra = Hour(Now())
    NuovoMinuti = Minute(Now()) + Minuti
    NuovoSecondi = Second(Now()) + Secondi
    Pausa = TimeSerial(NuovaOra, NuovoMinuti, NuovoSecondi)
    Application.Wait Pausa

End Sub

Public Sub Recalculate()

    'Procedure

    Application.Calculation = xlCalculationManual
    Calculate
 
End Sub

Public Sub BloccaXSecondi(Seconds As Integer)

    ' Wait a while blocking excel activity

    'Declarations

    Dim TheTime As Date

    'Procedure

    Let TheTime = Timer
    Do
    Loop Until (Timer - TheTime) >= Seconds

End Sub
