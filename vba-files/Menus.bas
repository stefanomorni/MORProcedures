Attribute VB_Name = "Menus"
Sub AggiungiComandoAMenu(Foglio As Object, Oggetto As String, Descrizione As String, PrimaDiMenu As Integer, MacroCodeName As String, Optional ByVal FaceId As Long = 0, Optional ByVal BeginGroup As Boolean = False)
    'oggetto puo essere p.es. 'List Range Popup' per le celle di una tabella o "Cell" per le celle fuori da una tabella

    Dim cmdBtn As CommandBarButton
    On Error Resume Next
    With Application
        .CommandBars(Oggetto).Controls(Descrizione).Delete
        Set cmdBtn = .CommandBars(Oggetto).Controls.Add(Before:=PrimaDiMenu, Temporary:=True)
    End With

    With cmdBtn
        .Caption = Descrizione
        If FaceId > 0 Then
            .Style = msoButtonIconAndCaption
            .FaceId = FaceId
        Else
            .Style = msoButtonCaption
        End If
        If BeginGroup Then
            .BeginGroup = True
        End If
        .OnAction = CodeName & MacroCodeName
    End With
    On Error GoTo 0
End Sub

Sub EliminaComandoDaMenu(Oggetto As String, Descrizione As String)
    'oggetto puo essere p.es. 'List Range Popup' per le ceeel di una tabella o "Cell" per le celle furi da una tabella

    On Error Resume Next
    With Application
        .CommandBars(Oggetto).Controls(Descrizione).Delete
    End With
    On Error GoTo 0
End Sub

Sub CreaMenuInStrumenti(Descrizione, Macro As String)
    Dim NewMenuItemMacro As String
    Dim NewItem As CommandBarButton
    Dim XLCommandBar As String
    Dim XLMenu As String
    Dim XLMenuItem As String
    Dim NewMenuItem As String
 
    XLCommandBar = "Worksheet Menu Bar"
    '   This code handles non-English versions of Excel
    '   in which the Tools menu has a different name
    XLMenu = Application.CommandBars(XLCommandBar).FindControl(msoControlPopup, 30007).Caption
    
    XLMenuItem = ""
    NewMenuItem = "&" & Descrizione & "..."
    NewMenuItemMacro = Macro
   
    '   Delete the current menu if it exists (just in case)
    On Error Resume Next
    Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls(XLMenuItem).Controls(NewMenuItem).Delete
    Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls(NewMenuItem).Delete
    On Error GoTo 0

    '   Create the new menu item
    If XLMenuItem = "" Then
        Set NewItem = Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls.Add
    Else
        Set NewItem = Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls(XLMenuItem).Controls.Add
    End If
   
    '   Specify the Caption and OnAction properties
    With NewItem
        .Caption = NewMenuItem
        .OnAction = NewMenuItemMacro
        .FaceId = 0                              'This is the image displayed next to the the menu item text
        .BeginGroup = True                       '   Add a separator bar before the menu item
    End With
    Exit Sub
    
    '   If an error occured, tell the user
    If Err <> 0 Then
        MsgBox "Si è verificato un errore nella procedura CreaMenuInStrumenti", vbInformation
    End If
End Sub

Sub EliminaMenuInStrumenti(Descrizione)
    '   This sub is executed when the workbook (or add-in) is closed.
    '   It simply removes the menu item
    
    Dim XLCommandBar As String
    Dim XLMenu As String
    Dim XLMenuItem As String
    Dim NewMenuItem As String
    
    XLCommandBar = "Worksheet Menu Bar"
    XLMenuItem = ""
    NewMenuItem = "&" & Descrizione & "..."

    '   This code handles non-English versions of Excel
    XLMenu = Application.CommandBars(XLCommandBar).FindControl(msoControlPopup, 30007).Caption
    On Error Resume Next
    Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls(XLMenuItem).Controls(NewMenuItem).Delete
    Application.CommandBars(XLCommandBar).Controls(XLMenu).Controls(NewMenuItem).Delete
End Sub

