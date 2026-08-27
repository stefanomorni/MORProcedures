Attribute VB_Name = "DDECalls"
'@Folder("MORProcedures.Modules")
Option Explicit
Option Base 1

'Declarations

Dim CanaleBlg
Dim CanaleFinXL

Public Sub AggregaRichiestaBlg(OggettoBlp, SerieTicker, SerieCampi, Optional CampiOverride, Optional ValoriOverride, Optional NomeVarRisposta, Optional Monitor)

    Dim OggettoBlp As BLP_DATA_CTRLLib.BlpData
    Dim Ticker, Campo, NoTrasmissione As Long

    OggettoBlp = New BlpData

    OggettoBlp.AutoRelease = False
    'Inizializza la serie di richieste dove 1 richiesta = 1Titolo,1 campo
    For Ticker = 1 To UBound(SerieTitoli)
        For Campo = 1 To UBound(SerieCampi)
            NoTrasmissione = NoTrasmissione + 1
            OggettoBlp.Subscribe SerieTitoli(Titolo), NoTrasmissione, SerieCampi(Campo), CampiOverride, ValoriOverride, NomeVarRisposta, Monitor
        Next Campo
    Next Ticker
    'invia le richieste aggregate
    OggettoBlp.Flush
End Sub

Public Sub ApriCanaleBlg()
        
    CanaleBlg = Application.DDEInitiate(App:="WinBlp", topic:="BBK")

End Sub

Public Sub ChiudiCanaleBlg()
        
    Application.DDETerminate CanaleBlg

End Sub

Public Sub StringaABlg(Stringa As String)
 
    'presuppone l'apertura del canale tramite Call ApriCanaleBlg
    Application.DDEExecute CanaleBlg, Stringa

End Sub

Public Sub SalvaSchermoBlg(NomeFile As String)

    StringaABlg ("<SAVE>")
    DoEvents
    Application.Wait (Now + TimeValue("0:00:03"))
    SendKeys NomeFile
    
End Sub
