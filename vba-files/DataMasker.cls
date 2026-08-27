VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataMasker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' DataMasker - High-Performance Deterministic In-Memory Range Masking Vault
' ==============================================================================
' - Deterministic pseudonymization: "CLI-" + 4-char hex hash
' - Sub-second execution using 2D Array memory blocks (Value2)
' - Non-contiguous range support iterating over rng.Areas
' - Vault persistence to %USERPROFILE%\.secrets\gp_data_mask_vault.json
' ==============================================================================

Private pForwardMap As Object ' Original -> Masked
Private pReverseMap As Object ' Masked -> Original
Private pVaultPath As String

Private Sub Class_Initialize()
    Set pForwardMap = CreateObject("Scripting.Dictionary")
    pForwardMap.CompareMode = vbBinaryCompare
    
    Set pReverseMap = CreateObject("Scripting.Dictionary")
    pReverseMap.CompareMode = vbBinaryCompare
    
    Dim userProfile As String
    userProfile = Environ("USERPROFILE")
    If Len(userProfile) = 0 Then
        userProfile = Environ("HOMEDRIVE") & Environ("HOMEPATH")
    End If
    
    Dim secretsDir As String
    secretsDir = userProfile & "\.secrets"
    EnsureDirectory secretsDir
    
    pVaultPath = secretsDir & "\gp_data_mask_vault.json"
    LoadVault
End Sub

Public Property Get VaultPath() As String
    VaultPath = pVaultPath
End Property

Public Function MaskSelection() As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range of cells to mask.", vbExclamation, "Data Masker"
        MaskSelection = 0
        Exit Function
    End If
    MaskRange Selection
    MaskSelection = Selection.Cells.count
End Function

Public Function UnmaskSelection() As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a valid range of cells to unmask.", vbExclamation, "Data Masker"
        UnmaskSelection = 0
        Exit Function
    End If
    UnmaskRange Selection
    UnmaskSelection = Selection.Cells.count
End Function

Public Sub MaskRange(ByVal targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevEnableEvents As Boolean
    
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim countModified As Long
    countModified = 0
    
    Dim area As Range
    For Each area In targetRng.Areas
        countModified = countModified + ProcessArea(area, True)
    Next area
    
    SaveVault

CleanUp:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    If Err.Number <> 0 Then
        MsgBox "Error during masking: " & Err.Description, vbCritical, "Data Masker"
    Else
        Application.StatusBar = "DataMasker: " & countModified & " cells masked."
    End If
End Sub

Public Sub UnmaskRange(ByVal targetRng As Range)
    If targetRng Is Nothing Then Exit Sub
    
    Dim prevScreenUpdating As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevEnableEvents As Boolean
    
    prevScreenUpdating = Application.ScreenUpdating
    prevCalculation = Application.Calculation
    prevEnableEvents = Application.EnableEvents
    
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim countModified As Long
    countModified = 0
    
    Dim area As Range
    For Each area In targetRng.Areas
        countModified = countModified + ProcessArea(area, False)
    Next area

CleanUp:
    Application.Calculation = prevCalculation
    Application.EnableEvents = prevEnableEvents
    Application.ScreenUpdating = prevScreenUpdating
    
    If Err.Number <> 0 Then
        MsgBox "Error during unmasking: " & Err.Description, vbCritical, "Data Masker"
    Else
        Application.StatusBar = "DataMasker: " & countModified & " cells restored."
    End If
End Sub

Private Function ProcessArea(ByVal areaRng As Range, ByVal isMasking As Boolean) As Long
    Dim cellCount As Long
    cellCount = areaRng.Cells.CountLarge
    
    Dim modifiedCount As Long
    modifiedCount = 0
    
    If cellCount = 1 Then
        Dim singleVal As Variant
        singleVal = areaRng.Value2
        If VarType(singleVal) = vbString And Len(singleVal) > 0 Then
            Dim singleStr As String
            singleStr = CStr(singleVal)
            If Left$(singleStr, 1) <> "=" Then
                If isMasking Then
                    Dim maskedSingle As String
                    maskedSingle = GetOrGenerateMask(singleStr)
                    If singleStr <> maskedSingle Then
                        areaRng.Value2 = maskedSingle
                        modifiedCount = 1
                    End If
                Else
                    If pReverseMap.Exists(singleStr) Then
                        areaRng.Value2 = pReverseMap(singleStr)
                        modifiedCount = 1
                    End If
                End If
            End If
        End If
        ProcessArea = modifiedCount
        Exit Function
    End If
    
    Dim dataArr As Variant
    dataArr = areaRng.Value2
    
    If Not IsArray(dataArr) Then
        ProcessArea = 0
        Exit Function
    End If
    
    Dim rMin As Long, rMax As Long, cMin As Long, cMax As Long
    rMin = LBound(dataArr, 1)
    rMax = UBound(dataArr, 1)
    cMin = LBound(dataArr, 2)
    cMax = UBound(dataArr, 2)
    
    Dim r As Long, c As Long
    Dim rawVal As Variant
    Dim rawStr As String
    Dim hasChanges As Boolean
    hasChanges = False
    
    For r = rMin To rMax
        For c = cMin To cMax
            rawVal = dataArr(r, c)
            If VarType(rawVal) = vbString And Len(rawVal) > 0 Then
                rawStr = CStr(rawVal)
                If Left$(rawStr, 1) <> "=" Then
                    If isMasking Then
                        Dim maskedVal As String
                        maskedVal = GetOrGenerateMask(rawStr)
                        If rawStr <> maskedVal Then
                            dataArr(r, c) = maskedVal
                            modifiedCount = modifiedCount + 1
                            hasChanges = True
                        End If
                    Else
                        If pReverseMap.Exists(rawStr) Then
                            dataArr(r, c) = pReverseMap(rawStr)
                            modifiedCount = modifiedCount + 1
                            hasChanges = True
                        End If
                    End If
                End If
            End If
        Next c
    Next r
    
    If hasChanges Then
        areaRng.Value2 = dataArr
    End If
    
    ProcessArea = modifiedCount
End Function

Private Function GetOrGenerateMask(ByVal originalStr As String) As String
    If pForwardMap.Exists(originalStr) Then
        GetOrGenerateMask = pForwardMap(originalStr)
        Exit Function
    End If
    
    ' Deterministic pseudonym: CLI- + 4 hex chars computed from CRC16/FNV-1a hash
    Dim hexHash As String
    hexHash = Compute4HexHash(originalStr)
    
    Dim candidate As String
    candidate = "CLI-" & hexHash
    
    ' Handle rare collision across distinct strings
    Dim collisionIndex As Long
    collisionIndex = 1
    While pReverseMap.Exists(candidate) And pReverseMap(candidate) <> originalStr
        candidate = "CLI-" & hexHash & "_" & CStr(collisionIndex)
        collisionIndex = collisionIndex + 1
    Wend
    
    pForwardMap(originalStr) = candidate
    pReverseMap(candidate) = originalStr
    GetOrGenerateMask = candidate
End Function

Private Function Compute4HexHash(ByVal s As String) As String
    ' CRC16 hash - pure Long arithmetic, no overflow risk
    Dim crc As Long
    crc = &HFFFF&

    Dim i As Long, j As Long, b As Long
    For i = 1 To Len(s)
        b = AscW(Mid$(s, i, 1)) And &HFF&
        crc = crc Xor b
        For j = 0 To 7
            If (crc And 1) <> 0 Then
                crc = (crc \ 2) And &H7FFF&
                crc = crc Xor &HA001&
            Else
                crc = (crc \ 2) And &H7FFF&
            End If
        Next j
    Next i

    Compute4HexHash = Right$("0000" & Hex$(crc), 4)
End Function

Private Sub LoadVault()
    pForwardMap.RemoveAll
    pReverseMap.RemoveAll
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(pVaultPath) Then Exit Sub
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(pVaultPath, 1, False, -2) ' -2 = TristateUseDefault / UTF-8 compatible ASCII
    Dim content As String
    If Not ts.AtEndOfStream Then
        content = ts.ReadAll
    End If
    ts.Close
    
    If Len(Trim$(content)) = 0 Then Exit Sub
    
    ParseVaultJson content
End Sub

Private Sub SaveVault()
    Dim jsonStr As String
    jsonStr = BuildVaultJson()
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.CreateTextFile(pVaultPath, True, False)
    ts.Write jsonStr
    ts.Close
End Sub

Private Function BuildVaultJson() As String
    Dim sb As String
    sb = "{" & vbCrLf & "  ""forward"": {" & vbCrLf
    
    Dim keys As Variant
    keys = pForwardMap.keys
    
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim k As String, v As String
        k = CStr(keys(i))
        v = CStr(pForwardMap(k))
        
        sb = sb & "    " & EscapeJsonString(k) & ": " & EscapeJsonString(v)
        If i < UBound(keys) Then
            sb = sb & "," & vbCrLf
        Else
            sb = sb & vbCrLf
        End If
    Next i
    
    sb = sb & "  }," & vbCrLf & "  ""reverse"": {" & vbCrLf
    
    Dim revKeys As Variant
    revKeys = pReverseMap.keys
    
    For i = LBound(revKeys) To UBound(revKeys)
        Dim rk As String, rv As String
        rk = CStr(revKeys(i))
        rv = CStr(pReverseMap(rk))
        
        sb = sb & "    " & EscapeJsonString(rk) & ": " & EscapeJsonString(rv)
        If i < UBound(revKeys) Then
            sb = sb & "," & vbCrLf
        Else
            sb = sb & vbCrLf
        End If
    Next i
    
    sb = sb & "  }" & vbCrLf & "}"
    BuildVaultJson = sb
End Function

Private Sub ParseVaultJson(ByVal json As String)
    ' Simple robust JSON parser for standard flat {"forward": {...}, "reverse": {...}} structure
    Dim forwardPos As Long, reversePos As Long
    forwardPos = InStr(1, json, """forward""", vbTextCompare)
    reversePos = InStr(1, json, """reverse""", vbTextCompare)
    
    If forwardPos = 0 Then Exit Sub
    
    Dim forwardEnd As Long
    If reversePos > forwardPos Then
        forwardEnd = reversePos
    Else
        forwardEnd = Len(json)
    End If
    
    ExtractPairs Mid$(json, forwardPos, forwardEnd - forwardPos), pForwardMap
    
    If reversePos > 0 Then
        ExtractPairs Mid$(json, reversePos), pReverseMap
    Else
        ' Rebuild reverse map from forward map if reverse block not present
        Dim k As Variant
        For Each k In pForwardMap.keys
            pReverseMap(pForwardMap(k)) = k
        Next k
    End If
End Sub

Private Sub ExtractPairs(ByVal section As String, ByRef dict As Object)
    Dim lines() As String
    lines = Split(section, vbCrLf)
    
    If UBound(lines) = 0 Then
        lines = Split(section, vbLf)
    End If
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim line As String
        line = Trim$(lines(i))
        
        ' Look for "key": "value"
        If Left$(line, 1) = """" Then
            Dim colonPos As Long
            colonPos = InStr(1, line, """:")
            If colonPos > 0 Then
                Dim keyRaw As String, valRaw As String
                keyRaw = Mid$(line, 1, colonPos)
                valRaw = Trim$(Mid$(line, colonPos + 2))
                
                ' Trim trailing comma
                If Right$(valRaw, 1) = "," Then valRaw = Trim$(Left$(valRaw, Len(valRaw) - 1))
                
                Dim parsedKey As String, parsedVal As String
                parsedKey = UnescapeJsonString(keyRaw)
                parsedVal = UnescapeJsonString(valRaw)
                
                If Len(parsedKey) > 0 Then
                    dict(parsedKey) = parsedVal
                End If
            End If
        End If
    Next i
End Sub

Private Function EscapeJsonString(ByVal s As String) As String
    Dim res As String
    res = s
    res = Replace(res, "\", "\\")
    res = Replace(res, """", "\""")
    res = Replace(res, vbCr, "\r")
    res = Replace(res, vbLf, "\n")
    res = Replace(res, vbTab, "\t")
    EscapeJsonString = """" & res & """"
End Function

Private Function UnescapeJsonString(ByVal s As String) As String
    Dim res As String
    res = Trim$(s)
    If Left$(res, 1) = """" And Right$(res, 1) = """" And Len(res) >= 2 Then
        res = Mid$(res, 2, Len(res) - 2)
    End If
    res = Replace(res, "\""", """")
    res = Replace(res, "\\", "\")
    res = Replace(res, "\r", vbCr)
    res = Replace(res, "\n", vbLf)
    res = Replace(res, "\t", vbTab)
    UnescapeJsonString = res
End Function

Private Sub EnsureDirectory(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        On Error Resume Next
        fso.CreateFolder folderPath
        On Error GoTo 0
    End If
End Sub
