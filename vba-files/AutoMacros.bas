Attribute VB_Name = "AutoMacros"
'@Folder("MORProcedures.Modules")
Option Explicit

Public Sub Auto_Open()
    RegisterShortcuts
End Sub

Public Sub Auto_Close()
    UnregisterShortcuts
End Sub

Public Sub RegisterShortcuts()
    On Error Resume Next
    Application.OnKey "^+M", "QuickMaskSelection"
    Application.OnKey "^+U", "QuickUnmaskSelection"
End Sub

Public Sub UnregisterShortcuts()
    On Error Resume Next
    Application.OnKey "^+M"
    Application.OnKey "^+U"
End Sub
