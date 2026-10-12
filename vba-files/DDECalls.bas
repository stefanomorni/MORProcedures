Attribute VB_Name = "DDECalls"
'@Folder("MORProcedures.Modules")
Option Explicit
Option Base 1

' Bloomberg DDE/COM integration — DEPRECATED 2026-10.
' Bloomberg terminal is no longer available on this workstation.
' All subs are retained as no-op stubs to avoid breaking any
' external callers that may still reference them.

Private Sub BloombergNotAvailable()
    Err.Raise 5, "DDECalls", "Bloomberg is no longer installed on this workstation."
End Sub

Public Sub AggregaRichiestaBlg(OggettoBlp As Variant, SerieTicker As Variant, SerieCampi As Variant, _
        Optional CampiOverride As Variant, Optional ValoriOverride As Variant, _
        Optional NomeVarRisposta As Variant, Optional Monitor As Variant)
    BloombergNotAvailable
End Sub

Public Sub ApriCanaleBlg()
    BloombergNotAvailable
End Sub

Public Sub ChiudiCanaleBlg()
    BloombergNotAvailable
End Sub

Public Sub StringaABlg(Stringa As String)
    BloombergNotAvailable
End Sub

Public Sub SalvaSchermoBlg(NomeFile As String)
    BloombergNotAvailable
End Sub
