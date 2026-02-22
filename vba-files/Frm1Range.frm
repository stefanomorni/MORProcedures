'@Folder("MORProcedures.Forms")
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm1Range 
   Caption         =   "Titolo"
   ClientHeight    =   2070
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5670
   OleObjectBlob   =   "Frm1Range.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm1Range"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Riferimento As String

'API function to enable/disable the Excel Window
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal bEnable As Long) As LongPtr

Dim mlHWnd As Long, mbModal As Boolean, mbDragDrop As Boolean

Private Sub UserForm_Activate()

    On Error Resume Next

    'Find the Excel main window
    mlHWnd = FindWindowA("XLMAIN", Application.Caption)
    mbDragDrop = Application.CellDragAndDrop     'Memorize the current state

    If CbxModeless.Value Then
        EnableWindow mlHWnd, 1                   'Enable the Window - makes the userform modeless
        'Disable Cell drag/drop, as it causes Excel 97 to GPF
        Application.CellDragAndDrop = False
    Else
        EnableWindow mlHWnd, 0                   'Disable the Window - makes the userform modal
    End If
End Sub

Private Sub CbOk_Click()
    Application.CellDragAndDrop = mbDragDrop
    Set RisultatoFinestra = Range(Riferimento)
    Unload Frm1Range
End Sub

Private Sub CbAnnulla_Click()
    Application.CellDragAndDrop = mbDragDrop
    Set RisultatoFinestra = Nothing
    Unload Frm1Range
End Sub
