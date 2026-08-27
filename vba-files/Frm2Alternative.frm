VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm2Alternative 
   Caption         =   "Finestra Informativa"
   ClientHeight    =   1785
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7536
   OleObjectBlob   =   "Frm2Alternative.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm2Alternative"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'API function to enable/disable the Excel Window
'Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As Long

Dim mlHWnd As Long, mbModal As Boolean, mbDragDrop As Boolean

Private Sub UserForm_Activate()

    On Error Resume Next

    'Find the Excel main window
    mlHWnd = FindWindowA("XLMAIN", Application.Caption)
    mbDragDrop = Application.CellDragAndDrop     'Memorize the current state

    If CbxModeless.value Then
        EnableWindow mlHWnd, 1                   'Enable the Window - makes the userform modeless
        'Disable Cell drag/drop, as it causes Excel 97 to GPF
        Application.CellDragAndDrop = False
    Else
        EnableWindow mlHWnd, 0                   'Disable the Window - makes the userform modal
    End If
End Sub

Private Sub CmdAlternativa1_Click()
    Application.CellDragAndDrop = mbDragDrop
    RisultatoFinestra = 1
    Unload Frm2Alternative
End Sub

Private Sub CmdAlternativa2_Click()
    Application.CellDragAndDrop = mbDragDrop
    RisultatoFinestra = 2
    Unload Frm2Alternative
End Sub

