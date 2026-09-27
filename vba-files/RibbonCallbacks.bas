Attribute VB_Name = "RibbonCallbacks"
'@Folder("MORProcedures.Ribbon")
Option Explicit

Private pRibbonUI As IRibbonUI

Public Sub Ribbon_OnLoad(ByVal ribbon As IRibbonUI)
    Set pRibbonUI = ribbon
End Sub

Public Sub OnRibbonMaskClick(ByVal control As IRibbonControl)
    QuickMaskSelection
End Sub

Public Sub OnRibbonUnmaskClick(ByVal control As IRibbonControl)
    QuickUnmaskSelection
End Sub

Public Sub QuickMaskSelection()
    Dim masker As DataMasker
    Set masker = New DataMasker
    Dim count As Long
    count = masker.MaskSelection()
    If count > 0 Then
        MsgBox "Successfully scrambled " & count & " cell value(s).", vbInformation, "Data Masking"
    End If
End Sub

Public Sub QuickUnmaskSelection()
    Dim masker As DataMasker
    Set masker = New DataMasker
    Dim count As Long
    count = masker.UnmaskSelection()
    If count > 0 Then
        MsgBox "Successfully restored " & count & " cell value(s).", vbInformation, "Data Masking"
    End If
End Sub