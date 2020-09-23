VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'*******************************************************
'*                                                     *
'*             BUGFULL FUNCTIONS v1.1                  *
'*          written by Chavdar Jordanov                *
'*                                                     *
'*******************************************************

Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" _
   Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
   ByVal lpszShortPath As String, ByVal cchBuffer As Long) _
   As Long

Private Declare Function GetLongPathName Lib "kernel32" Alias _
    "GetLongPathNameA" (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, ByVal cchBuffer As Long) _
    As Long
    
Private Declare Function LoadLibrary Lib "kernel32" _
  Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
  (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
  (ByVal hLibModule As Long) As Long
  
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Const MAX_PATH = 260


Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetDriveType32 Lib "kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long

Private Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFilename As String) As Long

Private Const cnstKeySeparator = "="
Private Const PunktString = ",./ +-*""'!@#$%^&*()_=?;:~`\|<>[]{}" + vbCrLf + vbTab

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOMOVE = &H2
Private Const fuFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

Public sWindowsDir As String, sSystemDir As String, sFontDir As String

Private Fmt As String

'crypt128 variables
Dim x1a0(9) As Long
Dim cle(17) As Long
Dim x1a2 As Long

Dim inter As Long, res As Long, ax As Long, bx As Long
Dim cx As Long, dx As Long, si As Long, tmp As Long
Dim iii As Long, C As Byte

'---
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2


Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Public Function BrowseFolder(ByVal handleWnd As Long, Optional ByVal szTitle = "Browse for folder") As String
'Opens a Browse Folders Dialog Box that displays the
'directories in your computer
Dim lpIDList As Long 'Declare Varibles
Dim sBuffer As String
Dim tBrowseInfo As BrowseInfo

With tBrowseInfo
   .hWndOwner = handleWnd 'Owner Form
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With

lpIDList = SHBrowseForFolder(tBrowseInfo)

If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   BrowseFolder = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End If
End Function

'============= INITIALIZATION ============
Private Sub Class_Initialize()
    GetSystemDirs
End Sub


'**********  STRING MANIPULATION FUNCTIONS **************

'This one capitalizes the first letter of the string
Public Function CapitalizeString(ByVal TheString As String) As String
    Dim NewString As String, X As Long, S As String, a$
    S = TrimString(TheString)
    If S = TheString Then
        a$ = UCase$(Left$(TheString, 1))
        CapitalizeString = a$ + Mid$(TheString, 2)
    Else
        X = InStr(1, TheString, S, vbBinaryCompare)
        a$ = UCase$(Left$(S, 1)) + Mid$(S, 2)
        CapitalizeString = Left$(TheString, X - 1) + a$ + Mid$(TheString, X + Len(S))
    End If
End Function

'This replaces the OldValue in TheString with a NewValue
Public Function ReplaceString(ByVal TheString As String, ByVal OldValue As String, ByVal NewValue As String, iCompare, MatchCase As Boolean, WholeWord As Boolean) As String
    Dim X As Long, s1 As String, s2 As String, s3 As String
    X = 0
    Do
        X = InStr(X + 1, TheString, OldValue, iCompare)
        If X = 0 Then Exit Do
        s1 = Left$(TheString, X - 1)
        s2 = Mid$(TheString, X + Len(OldValue))
        If WholeWord Then
            If InStr(1, PunktString, Right$(s1, 1), vbBinaryCompare) = 0 And InStr(1, PunktString, Left$(s2, 1)) = 0 Then GoTo 10
        End If
        If MatchCase Then
            s3 = TrimString(Mid$(TheString, X, Len(OldValue)))
            If UCase(s3) = s3 Then
                NewValue = UCase(NewValue)
            ElseIf LCase(s3) = s3 Then
                NewValue = LCase(NewValue)
            ElseIf CapitalizeString(s3) = s3 Then
                NewValue = CapitalizeString(NewValue)
            End If
        End If
        TheString = s1 + NewValue + s2
10
    Loop
    ReplaceString = TheString
End Function

'This one APPENDS up to N spaces to the string
Public Function FillLineWithSpaces(TheString As String, NumberOfSpaces As Integer) As String
    Dim L As Long
    L = Len(TheString)
    FillLineWithSpaces = TheString + Space(NumberOfSpaces - L)
End Function

'This reverses the string
Public Function ReverseString(ByVal TheString As String) As String
    Dim L As String, z As String
    Dim i As Long
    L = Len(TheString)
    z = ""
    For i = L To 1 Step -1
        z = z + Mid$(TheString, i, 1)
    Next i
    ReverseString = z
End Function

Public Function GetDirFromFileName(ByVal FileName As String, Optional Separator = "\") As String
    Dim i As Long
    i = InStrRev(FileName, Separator)
    If i <> 0 Then
        GetDirFromFileName = Left(FileName, i - 1)
    Else
        GetDirFromFileName = FileName
    End If
    
End Function


Public Function LastInstr(ByVal String1 As String, String2 As String) As Long
    LastInstr = InStrRev(String1, String2, -1, vbTextCompare)
End Function

Public Function leftPart(ByVal TheString As String, Separator As String, ByVal FromEnd As Boolean) As String
    Dim X As Long
    If Not FromEnd Then X = InStr(1, TheString, Separator, vbTextCompare) Else X = LastInstr(TheString, Separator)
    If X = 0 Then GoTo 10
    leftPart = Left$(TheString, X - 1)
10
End Function

Public Function rightPart(ByVal TheString As String, Separator As String, ByVal FromEnd As Boolean) As String
    Dim X As Long
    If Not FromEnd Then X = InStr(1, TheString, Separator, vbTextCompare) Else X = LastInstr(TheString, Separator)
    If X = 0 Then GoTo 10
    rightPart = Mid$(TheString, X + 1)
10
End Function

Public Function GetDriveLetter(sPath As String) As String
    Dim L As Long, i As Long, C As String
    If Mid$(sPath, 2, 1) = ":" Then
        GetDriveLetter = Left$(sPath, 1)
    Else
        Dim SlashCnt As Integer
        SlashCnt = 0
        L = Len(sPath)
        For i = 1 To L
            C = Mid$(sPath, i, 1)
            If C = "\" Then SlashCnt = SlashCnt + 1
            If SlashCnt = 3 Then
                GetDriveLetter = UCase$(Mid$(sPath, i + 1, 1))
                Exit For
            End If
        Next i
    End If
End Function

Public Function UpDir(sDir As String) As String 'returns the parent directory of sDir or empty string
    Dim i, S, L
    L = Len(sDir)
    For i = L To 1 Step -1
        If i < L And Mid$(sDir, i, 1) = "\" Then
            UpDir = Left$(sDir, i)
            Exit Function
        End If
    Next i
End Function

Public Function TrimString(ByVal TheString As String) As String 'Trims the string from spaces,punctuation, etc
    Dim i, L, C As String, NewString As String, FirstMet As Long, LastMet As Long
    
    L = Len(TheString)
    FirstMet = False
    For i = 1 To L
        C = Mid$(TheString, i, 1)
        If InStr(1, PunktString, C, vbBinaryCompare) = 0 Then
            FirstMet = i
            Exit For
        End If
    Next i
    For i = L To 1 Step -1
        C = Mid$(TheString, i, 1)
        If InStr(1, PunktString, C, vbBinaryCompare) = 0 Then
            LastMet = i
            Exit For
        End If
    Next i
    If FirstMet < LastMet Then
        TrimString = Mid$(TheString, FirstMet, LastMet - FirstMet + 1)
    End If
End Function
'==========================================================================
'= FILE HANDLING FUNCTIONS =

Public Function ExcludePath(ByVal RootDir, ByVal FileDir) As String
   Dim xz As Long
    If FileDir = "" Or RootDir = "" Then
        ExcludePath = FileDir
    Else
        xz = InStr(1, FileDir, RootDir, vbTextCompare)
        If xz = 0 Then
            ExcludePath = FileDir
        Else
            ExcludePath = Mid(FileDir, xz + Len(RootDir))
        End If
    End If
End Function

Public Function DateToDbl(ByVal dDate) As Double
    DateToDbl = CDbl(dDate)
End Function

Public Function PathToNetwork(ByVal sPcName, ByVal sPath) As String
    If sPcName = "" Then PathToNetwork = sPath Else PathToNetwork = "\\" + sPcName + "\" + Left(sPath, 1) + Mid(sPath, 3)
End Function

Public Function GetSystemDirs()
    Dim X
    Const nBuffer = 255
    sWindowsDir = Space$(nBuffer)
    X = GetWindowsDirectory(sWindowsDir, nBuffer)
    sWindowsDir = Trim(sWindowsDir)
    sWindowsDir = Left(sWindowsDir, Len(sWindowsDir) - 1) + "\"
    
    sSystemDir = Space$(nBuffer)
    X = GetSystemDirectory(sSystemDir, nBuffer)
    sSystemDir = Trim(sSystemDir)
    sSystemDir = Left(sSystemDir, Len(sSystemDir) - 1) + "\"
    
    sFontDir = sWindowsDir + "Fonts\"
End Function


Public Sub CreateDir(sPath As String)
    Dim x1, x2, NewDir As String
    sPath = BF.ToPath(sPath)
    x1 = InStr(sPath, "\\")
    x1 = InStr(x1 + 2, sPath, "\", vbBinaryCompare)
    Do
      x2 = InStr(x1 + 1, sPath, "\", vbBinaryCompare)
      If x2 = 0 Then Exit Sub
      NewDir = Left(sPath, x2)
      If Dir(NewDir, vbDirectory) = "" Then MkDir NewDir
      x1 = x2
    Loop Until x2 = Len(sPath)
    
End Sub

Public Function TypeOfDrive(sDriveLetter As String) As String
    Dim sDrive As String, lRet As Long
    sDrive = sDriveLetter + ":\"
    lRet = GetDriveType32(sDrive)
    Select Case lRet
        Case 3
            TypeOfDrive = "fixed"
        Case 5
            TypeOfDrive = "cd"
        Case 2
            TypeOfDrive = "removable"
        Case 4
            TypeOfDrive = "remote"
        Case Else
            TypeOfDrive = "unknown"
    End Select
End Function

Public Function GetCDRoms() As String
    Const Drives = "defghijklmnopqrstuvwxyz"
    Dim L, sCDRoms As String, i, cDrv As String
    L = Len(Drives)
    sCDRoms = ""
    For i = 1 To L
        cDrv = Mid(Drives, i, 1)
        If TypeOfDrive(cDrv) = "cd" Then
            sCDRoms = sCDRoms + cDrv
        End If
    Next i
    GetCDRoms = sCDRoms
End Function

Public Function ToPath(ByVal sPath As String) As String
    If sPath <> "" Then If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    ToPath = sPath
End Function


'=====================================================================
'=         FORM HANDLING FUNCTIONS

Public Sub AlwaysOnTop(FormHwnd As Long, ByVal TopMost As Boolean)
    If TopMost = True Then
      SetWindowPos FormHwnd, -1, 0, 0, 0, 0, fuFlags
    Else
        ' Turn off the TopMost attribute.
      SetWindowPos FormHwnd, -2, 0, 0, 0, 0, fuFlags
    End If
End Sub

Public Sub CenterForm(TheForm As Object)
    With TheForm
        .Left = (Screen.Width - .Width) / 2
        .Height = (Screen.Height - .Height) / 2
    End With
End Sub


'=====================================================================
'=        CRYPTING FUNCTIONS

Public Function Crypt(inp As String, key As String) As String
    Dim Sbox(0 To 255) As Long
    Dim Sbox2(0 To 255) As Long
    Dim j As Long, i As Long, t As Double
    Dim K As Long, temp As Long, Outp As String
    Dim X
    
    For i = 0 To 255 'Create SBox #1
                    Sbox(i) = i  'and fill with
    Next i           'successive values

     j = 1
    For i = 0 To 255                   'Create SBox #2
                    If j > Len(key) Then j = 1     'And fill with key
                    Sbox2(i) = Asc(Mid(key, j, 1)) 'data, repeatedly
             j = j + 1                      'until the SBox is
    Next i                             'full

     j = 0 'Initialize j
     For i = 0 To 255                         'Scramble SBox #1
             j = (j + Sbox(i) + Sbox2(i)) Mod 256 'with data from
            temp = Sbox(i)                       'SBox #2
                    Sbox(i) = Sbox(j)
                    Sbox(j) = temp
    Next i

     i = 0 'Initialize i
     j = 0 'Initialize j
    For X = 1 To Len(inp) 'Process the data passed on to us
             i = (i + 1) Mod 256 'Increment i
             j = (j + Sbox(i)) Mod 256 'Increment j
             temp = Sbox(i)    'Scramble SBox #1
                    Sbox(i) = Sbox(j) 'further so that the encryptor
                    Sbox(j) = temp    'will never repeat itself
             t = (Sbox(i) + Sbox(j)) Mod 256 'Get ready to create "random" byte
                    K = Sbox(t) 'Get "random" byte

                    Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K) 'Xor the data with the "random" byte
    Next X
    Crypt = Outp 'Return the Output Data
    'Debug.Print Len(Outp)
End Function

Sub Assemble()
    x1a0(0) = ((cle(1) * 256) + cle(2)) Mod 65536
    Code
    inter = res
    
    x1a0(1) = x1a0(0) Xor ((cle(3) * 256) + cle(4))
    Code
    inter = inter Xor res
    
    
    x1a0(2) = x1a0(1) Xor ((cle(5) * 256) + cle(6))
    Code
    inter = inter Xor res
    
    x1a0(3) = x1a0(2) Xor ((cle(7) * 256) + cle(8))
    Code
    inter = inter Xor res
    
    x1a0(4) = x1a0(3) Xor ((cle(9) * 256) + cle(10))
    Code
    inter = inter Xor res
    
    x1a0(5) = x1a0(4) Xor ((cle(11) * 256) + cle(12))
    Code
    inter = inter Xor res
    
    x1a0(6) = x1a0(5) Xor ((cle(13) * 256) + cle(14))
    Code
    inter = inter Xor res
    
    x1a0(7) = x1a0(6) Xor ((cle(15) * 256) + cle(16))
    Code
    inter = inter Xor res
    iii = 0
End Sub

Sub Code()
    dx = (x1a2 + iii) Mod 65536
    ax = x1a0(iii)
    cx = &H15A
    bx = &H4E35
    
    tmp = ax
    ax = si
    si = tmp
    
    tmp = ax
    ax = dx
    dx = tmp
    
    If (ax <> 0) Then
        ax = (ax * bx) Mod 65536
    End If
    
    tmp = ax
    ax = cx
    cx = tmp
    
    If (ax <> 0) Then
        ax = (ax * si) Mod 65536
        cx = (ax + cx) Mod 65536
    End If
    
    tmp = ax
    ax = si
    si = tmp
    ax = (ax * bx) Mod 65536
    dx = (cx + dx) Mod 65536
    
    ax = ax + 1
    
    x1a2 = dx
    x1a0(iii) = ax
    
    res = ax Xor dx
    iii = iii + 1

End Sub

Public Function Crypt128(ByVal inp As String, ByVal key As String, sourceLen As Long) As String
    Dim fois As Long
    Dim champ1 As String
    Dim lngchamp1 As Long
    Dim cfc, cfd
    Dim compte As Byte
    Dim d, e
    Crypt128 = ""
    si = 0
    x1a2 = 0
    iii = 0
    
    For fois = 1 To 16
        cle(fois) = 0
    Next fois
    champ1 = key
    lngchamp1 = Len(champ1)
    For fois = 1 To lngchamp1
        cle(fois) = Asc(Mid(champ1, fois, 1))
    Next fois
    champ1 = inp
    lngchamp1 = Len(champ1)
    For fois = 1 To lngchamp1
        C = Asc(Mid(champ1, fois, 1))
        Assemble
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
        For compte = 1 To 16
            cle(compte) = cle(compte) Xor C
        Next compte
        C = C Xor (cfc Xor cfd)
        d = (((C / 16) * 16) - (C Mod 16)) / 16
        e = C Mod 16
        Crypt128 = Crypt128 + Chr$(&H61 + d) ' d+&h61 give one letter range from a to p for the 4 high bits of c
        Crypt128 = Crypt128 + Chr$(&H61 + e) ' e+&h61 give one letter range from a to p for the 4 low bits of c
    Next fois
End Function

Public Function Decrypt128(sText As String, key As String)
    Dim fois As Long
    Dim champ As String, champ1 As String
    Dim lngchamp1 As Long
    Dim d As Integer, e As Integer
    Dim cfc
    Dim cfd
    Dim compte As Long
    
    Decrypt128 = ""
    si = 0
    x1a2 = 0
    iii = 0
    For fois = 1 To 16
        cle(fois) = 0
    Next
    champ1 = key
    lngchamp1 = Len(champ1)
    For fois = 1 To lngchamp1
        cle(fois) = Asc(Mid(champ1, fois, 1))
    Next
    champ1 = sText
    lngchamp1 = Len(champ1)
    For fois = 1 To lngchamp1
        d = Asc(Mid(champ1, fois, 1))
        If (d - &H61) >= 0 Then
            d = d - &H61
            If (d >= 0) And (d <= 15) Then
                d = d * 16
            End If
        End If
        If (fois <> lngchamp1) Then
            fois = fois + 1
        End If
        e = Asc(Mid(champ1, fois, 1))
        If (e - &H61) >= 0 Then
            e = e - &H61
            If (e >= 0) And (e <= 15) Then
                C = d + e
            End If
        End If
        
        Assemble
        cfc = (((inter / 256) * 256) - (inter Mod 256)) / 256
        cfd = inter Mod 256
        C = C Xor (cfc Xor cfd)
        For compte = 1 To 16
            cle(compte) = cle(compte) Xor C
        Next compte
        Decrypt128 = Decrypt128 + Chr(C)
    Next fois
End Function


'======================== INI FILES HANDLING FUNCTIONS =====================

Public Function GetProfileString(ByVal sFile As String, ByVal sSection As String, ByVal sKey As String, Optional ByVal DefaultValue = "") As String
    Dim sTmp As String, X As Long
    Const StringSize = 1024
    sTmp = Space$(StringSize)
    X = GetPrivateProfileString(sSection, sKey, "", sTmp, StringSize, sFile)
    sTmp = Trim$(sTmp)
    sTmp = Left(sTmp, Len(sTmp) - 1)
    If sTmp = "" Then sTmp = DefaultValue
    GetProfileString = sTmp
End Function

Public Sub WriteProfileString(ByVal FileName As String, ByVal SeparatorName As String, ByVal KeyName As String, ByVal KeyValue)
    Dim X
    X = WritePrivateProfileString(SeparatorName, KeyName, CStr(KeyValue), FileName)
End Sub
Public Function ToSection(ByVal sSectionName) As String
    If sSectionName <> "" Then
        ToSection = "[" + sSectionName + "]"
    Else
        ToSection = ""
    End If
End Function

Public Function ToKey(ByVal sKeyName, ByVal sKeyValue)
    
    If sKeyName <> "" Then
        ToKey = sKeyName + cnstKeySeparator + sKeyValue
    End If
End Function

Public Sub AddSection(ByVal sFileName As String, ByVal sSectionName As String)
    Dim nFree
    nFree = FreeFile
    On Error GoTo 100
    Open sFileName For Append As #nFree
    Print #nFree,
    Print #nFree, "[" + sSectionName + "]"
5
    Close #nFree
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 5
End Sub

Public Sub RemoveSection(ByVal sFileName As String, ByVal sSectionName As String)
    Dim nFree1, nFree2, bFlag As Boolean
    Dim sTmpName As String, sLine As String
    On Error GoTo 100
    sTmpName = BF.ToPath(App.Path) + "tpm543.tmp"
    
    nFree1 = FreeFile
    Open sFileName For Input As #nFree1
    nFree2 = FreeFile
    Open sTmpName For Output As #nFree2
    While Not EOF(nFree1)
        Line Input #nFree1, sLine
        If Trim(UCase(sLine)) = "[" + UCase(sSectionName) + "]" Then bFlag = True: GoTo 5
        If Left$(sLine, 1) = "[" And bFlag Then bFlag = False
        If Not bFlag Then Print #nFree2, sLine
5
    Wend
    Close #nFree1
    Close #nFree2
    Kill sFileName
    Name sTmpName As sFileName
10
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

Public Function GetSection(sFileName As String, sSectionName As String, sKeys() As String) As Integer
    Dim nFree, sLine As String, bFlag As Boolean, Cnt As Integer
    nFree = FreeFile
    Open sFileName For Input As #nFree
    While Not EOF(nFree)
        Line Input #nFree, sLine
        If Trim(sLine) = "[" + sSectionName + "]" Then bFlag = True: GoTo 5
        If Left$(sLine, 1) = "[" And bFlag Then bFlag = False
        If bFlag Then
            If InStr(sLine, cnstKeySeparator) > 0 Then
                Cnt = Cnt + 1
                ReDim Preserve sKeys(1 To Cnt)
                sKeys(Cnt) = sLine
            End If
        End If
5
    Wend
    GetSection = Cnt
End Function


'======================== CALCULATION FUNCTIONS =====================

Public Function Calc(ByVal Ex As String) As Double
    Dim sX As String
    Dim x1, i
    Dim z As String
    Dim cn As String
    Dim CFM$
    Dim X
    
    ' remove spaces from Ex
    sX = Condense(Ex)
    x1 = 0
    Do
5       For i = Len(sX) To 1 Step -1
            cn = Mid$(sX, i, 1)
            If cn = ")" Then x1 = i
            If cn = "(" Then
                z = SubCalc(Mid$(sX, i + 1, x1 - i - 1))
                sX = Left$(sX, i - 1) & Format(z, CFM$) & Mid$(sX, x1 + 1)
                x1 = 0
                GoTo 5
            End If
        Next i
        If x1 <> 0 Then X = 1 / 0: Exit Function
        Exit Do
    Loop
    Calc = SubCalc(sX)
End Function

Public Function Condense(ByVal S As String) As String
    Dim z As String
    Dim i As Long
    Dim ss As String
    
    z = ""
    For i = 1 To Len(S)
        ss = Mid$(S, i, 1)
        If ss > " " Then z = z + ss
    Next i
    Condense = z
End Function

Public Function SubCalc(ByVal S As String) As Double
    Dim Ns(1 To 200) As Double, Sg(1 To 200) As String
    Dim L As Long
    Dim i, oldi, j
    Dim Char As String * 1
    Dim z As String * 1
    Dim aa
    Dim cn
    Dim res As Integer
    Dim Znak
    
    S = S + "="
    L = Len(S)
    cn = 0:  oldi = 1
    For i = 1 To L
       Char = Mid$(S, i, 1)
       If InStr(1, "+-*/^=", Char, vbTextCompare) <> 0 Then
          cn = cn + 1
          z = Mid$(S, oldi, i - 1)
          If Left$(z, 1) = "!" Then z = "-" + Mid$(z, 2)
          Ns(cn) = Val(z)
          Sg(cn) = Char
          oldi = i + 1
        Else
            If InStr(1, "0123456789,.!", Char, vbBinaryCompare) = 0 Then
                aa = 1 / 0
                'MsgBox "Error in expression": Exit Function
            End If
        End If
        'Check for errors
    Next i
    res = 0
    Znak = Array("^", "/", "*", "+", "-", "=")
    Do
        i = 1
        Do
            If Sg(i) = Znak(res) Then
                Select Case Sg(i)
                    Case "^"
                        Ns(i) = Ns(i) ^ Ns(i + 1)
                    Case "*"
                        Ns(i) = Ns(i) * Ns(i + 1)
                    Case "/"
                        Ns(i) = Ns(i) / Ns(i + 1)
                    Case "+"
                        Ns(i) = Ns(i) + Ns(i + 1)
                    Case "-"
                        Ns(i) = Ns(i) - Ns(i + 1)
                End Select
                For j = i + 1 To cn
                    Ns(j) = Ns(j + 1)
                    Sg(j - 1) = Sg(j)
                Next j
                cn = cn - 1: If cn = 0 Then Exit Do
                i = 1
            Else
                i = i + 1: If i > cn Then Exit Do
            End If
        Loop
        res = res + 1: If res > 5 Then Exit Do
    Loop
    SubCalc = Ns(1)
End Function

Public Function MyVal(ByVal S) As Double
    Dim C As String, z As String, i As Long
    Dim L As Long
    L = Len(S)
    z = ""
    For i = 1 To L
        C = Mid$(S, i, 1)
        If InStr("0123456789.", C) <> 0 Then z = z + C
    Next i
    If z = "" Then MyVal = 0 Else MyVal = CDbl(z)
End Function


'=================== ARRAY HANDLING FUNCTIONS ===============

Public Sub BubbleSort(ByRef lArray())
    Dim i, bFlag As Boolean, C
    Dim L, u
    L = LBound(lArray)
    u = UBound(lArray) - 1
    Do
        bFlag = False
        For i = L To u
            If lArray(i) > lArray(i + 1) Then
                C = lArray(i + 1)
                lArray(i + 1) = lArray(i)
                lArray(i) = C
                bFlag = True
            End If
        Next i
        If Not bFlag Then Exit Do
    Loop
End Sub

Public Function GetLongFileName(ByVal FullPath As String) _
   As String

'*****************************************
'USAGE: Convert short (8.3) file name to long file name
'INPUT: FULL PATH OF A SHORT FILE NAME
'RETURNS: LONG FILE NAME:
'EXAMPLE: dim sLongFile as String
'         sLongFile = GetLongFileName("C\:MyShor~1.txt")
'NOTES: ONLY WORKS ON WIN 98 and WIN 2000.  WILL RETURN
'       EMPTY STRING ELSEWHERE
'***********************************************************

    Dim lLen As Long
    Dim sBuffer As String
    
    'Function only available on '98 and 2000,
    'so we check to see if it's available before proceeding
    
    If Not APIFunctionPresent("GetLongPathNameA", "kernel32") _
       Then Exit Function
    
    sBuffer = String$(MAX_PATH, 0)
    lLen = GetLongPathName(FullPath, sBuffer, Len(sBuffer))
    If lLen > 0 And Err.Number = 0 Then
        GetLongFileName = Left$(sBuffer, lLen)
    End If
End Function

Private Function APIFunctionPresent(ByVal FunctionName _
   As String, ByVal DllName As String) As Boolean

   'http://www.freevbcode.com/ShowCode.Asp?ID=429

    Dim lHandle As Long
    Dim lAddr  As Long

    lHandle = LoadLibrary(DllName)
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    
    APIFunctionPresent = (lAddr <> 0)

End Function


Public Function GetShortFileName(ByVal FullPath As String) _
  As String

    'PURPOSE: Returns DOS File Name (8.3 Format) Give
    'FullPath for long file name
    
    'PARAMETERS: FullPath: Full Path of Original File
    
    'RETURNS: 8.3 FileName, or "" if FullPath doesn't
    '         exist or file fails for other reasons
    
    'EXAMPLE:
    'Debug.Print _
    '  GetShortFileName("C:\My Documents\My Very Long File Name.doc")
    
    'If file exists, will display C:\MYDOCU~1\MYVERY~1.DOC
    'in the debug window
    
    Dim lAns As Long
    Dim sAns As String
    Dim iLen As Integer
       
    On Error Resume Next
    
    'this function doesn't work if the file doesn't exist
    If Dir(FullPath) = "" Then Exit Function
    
    sAns = Space(255)
    lAns = GetShortPathName(FullPath, sAns, 255)
    GetShortFileName = Left(sAns, lAns)
    
End Function

Public Function RemoveExtension(sFileName As String) As String
    Dim X As Long
    X = InStrRev(sFileName, ".")
    If X <> 0 Then
        RemoveExtension = Left(sFileName, X - 1)
    Else
        RemoveExtension = sFileName
    End If
End Function

Public Function IsDirectory(sPath As String) As Boolean
    On Error Resume Next
    IsDirectory = GetAttr(sPath) And vbDirectory
End Function

Public Function GetFileNameFromDir(ByVal DirPath As String) As String
    Dim L As Long, i As Long
    L = Len(DirPath)
    i = InStrRev(DirPath, "\")
    If i = 0 Then
        GetFileNameFromDir = DirPath
    Else
        GetFileNameFromDir = Mid(DirPath, i + 1)
    End If
End Function

Public Function ShellOpen(lHWND As Long, sFileName As String) As Boolean
    Dim lRet As Long
    lRet = ShellExecute(lHWND, "open", sFileName, "", "", 3)
    ShellOpen = lRet > 32
End Function
