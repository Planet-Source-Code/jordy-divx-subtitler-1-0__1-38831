Attribute VB_Name = "modSubtitle"
'*****************************************************************
'*              Subtitle processing subroutines                  *
'*              written by Chavdar Yordanov, 04.2001             *
'*              Email: chavdar_jordanov@yahoo.com                *
'*              Please, don't remove this title!                 *
'*****************************************************************
Option Explicit

'A custom type to use with subtitles
Public Type Subtitles
    StartFrame As Long  'starting frame of the title
    EndFrame As Long    'ending frame of the title
    Format As String    'title format (not in use)
    Text As String      'title body
End Type

Public BF As New clsFunctions 'an object containing some useful functions

Public Title() As Subtitles   'the main array containg the titles
Public Mark(1 To 7) As Long   '7 movie position markers
Public LineCnt As Long        'contains the number of subtitles in Title()
Public lSelStart As Long      'selection start and end indexes
Public lSelEnd As Long
Public gFPS As Single         'the movie's frame rate
Public bTitleStart As Boolean
Public lUpdateTime As Integer  'subtitles update interval (when playing the movie)
Public lMeasureTime As Integer 'time needed to measure the frame rate
Public sSubFileName As String  'subtitle file name
Public sAviFile As String      'movie file name
Public bAutoMeasure As Boolean 'auto-measure FPS on movie load
Public lRowFound As Long       'used with Find&Replace dialog
Public bChanged As Boolean     'flag that the file has been changed and should be saved

'Some constants
Public Const cst_AppName = "Subtitler"
Public Const cst_Settings = "Settings"
Public Const ckey_UpdateTime = "UpdateTime"
Public Const ckey_MeasureTime = "MeasureTime"
Public Const ckey_AutoMeasure = "AutoMeasure"

'--- finds the starting position of the first subtitle in a MDVD file ---
Function GetFirstIndex(ByRef S As String) As Long
    Dim X, x2
    Dim NewFrames As String, Start
    X = InStr(1, S, "{", vbBinaryCompare)
    x2 = InStr(X, S, "}", vbBinaryCompare)
    GetFirstIndex = Val(Mid(S, X + 1, x2 - X - 1))
End Function

'--- Reads and loads the subtitle file into an array;
'    tries to recognize the file format automaticaly ---
Public Function GetTitles(sFileName As String, ByRef tTitle() As Subtitles) As Long
    Dim H, G
    Dim sStart, sEnd, sText
    Dim S As String
    Dim i As Long
    Dim L As Long, lLineCnt As Long, t
    Dim sExt As String
    Dim RowCnt As Integer
    Dim bWasEmpty As Boolean
    
    On Error GoTo 0
    'read the file into a string
    S = ReadFile(sFileName)
    'split the string at the line breaks
    H = Split(S, vbCrLf)
    'count the lines
    L = UBound(H)
    'attempt to recognize the file format:
    If InStr(S, "}{") Then 'MDVD Format
        lLineCnt = 0
        ReDim tTitle(1 To L + 1)
        For i = 0 To L
            If H(i) <> "" Then
                If Left(H(i), 1) = Chr(10) Then H(i) = Mid(H(i), 2)
                'split the line at }'s :
                G = Split(H(i), "}", 3)
                If UBound(G) > 1 Then
                    If Trim(G(2)) <> "" Then
                        lLineCnt = lLineCnt + 1
                        tTitle(lLineCnt).StartFrame = Val(Mid(G(0), 2))
                        tTitle(lLineCnt).EndFrame = Val(Mid(G(1), 2))
                        tTitle(lLineCnt).Text = (G(2))
                    Else
                        'Stop
                    End If
                End If
            End If
        Next i
    ElseIf InStr(S, "-->") Then 'SubRip format
        RowCnt = 0
        For i = 0 To L
            RowCnt = RowCnt + 1 'line counter for a single title (usually consists of 3-5 lines)
            Select Case RowCnt
            Case 1 'the first row contains the title counter; we don't need it
                'so do nothing
            Case 2 'The second row contains the time code
                t = Split(H(i), " ")
                lLineCnt = lLineCnt + 1
                ReDim Preserve tTitle(1 To lLineCnt)
                'recalculate the time code into frames
                tTitle(lLineCnt).StartFrame = ConvertFromTime(t(0))
                tTitle(lLineCnt).EndFrame = ConvertFromTime(t(2))
            Case Else 'all the rest is the title text
                If IsNumeric(H(i)) And bWasEmpty Then 'we encounter the title count line
                    'so remember the last title text
                    tTitle(lLineCnt).Text = Mid(sText, 2)
                    'and reset the variables
                    RowCnt = 1
                    sText = ""
                    bWasEmpty = False
                ElseIf Trim(H(i)) <> "" Then 'add the title text to a buffer string
                    sText = sText + "|" + H(i)
                Else
                    bWasEmpty = True 'flag that we had an empty row and are looking for the next title
                End If
            End Select
        Next i
    Else
        MsgBox "Unrecognized subtitles format.", vbCritical
    End If
    S = ""
    Erase H
    ReDim Preserve tTitle(1 To lLineCnt)
    GetTitles = lLineCnt
End Function

'--- Saves/appends the titles to a file according to the format (defined by the extension) ---
Public Function SaveTitles(sFileName As String, tTitle() As Subtitles, Optional bAppend = False) As Boolean
    Dim i As Long
    Dim L As Long, F As Long
    Dim sExt As String
    sExt = LCase(Right(sFileName, 3))
    L = UBound(tTitle)
    F = FreeFile
    On Error GoTo 100
    Select Case sExt
    Case "srt"
        If gFPS = 0 Then
            MsgBox "Please, measure the movie FPS before saving in this format!", vbExclamation
            Exit Function
        End If
        If bAppend Then
            MsgBox "Sorry, Subtitler can not append to this format.", vbCritical
            Exit Function
        End If
        Screen.MousePointer = 11
        Open sFileName For Output As #F
        For i = 1 To L
            Print #F, CStr(i)
            Print #F, Convert2Time(tTitle(i).StartFrame) + " --> " + Convert2Time(tTitle(i).EndFrame)
            Print #F, Replace(tTitle(i).Text, "|", vbCrLf)
            Print #F,
        Next i
        Close #F
    Case Else
        Screen.MousePointer = 11
        If bAppend Then
            Open sFileName For Append As #F
        Else
            Open sFileName For Output As #F
        End If
        For i = 1 To L
            Print #F, "{" + CStr(tTitle(i).StartFrame) + "}{" + CStr(tTitle(i).EndFrame) + "}" + CStr(tTitle(i).Text)
        Next i
        Close #F
    End Select
    SaveTitles = True
10
    Screen.MousePointer = 0
    Exit Function
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Function

Sub ChangeIndex(ByVal StartTitle As Long, ByVal bIsStart As Boolean, ByVal Change As Double)
    Dim i As Long
    Screen.MousePointer = 11
    For i = StartTitle To LineCnt
        If (Not bIsStart And i > StartTitle) Or bIsStart Then Title(i).StartFrame = Title(i).StartFrame + Change
        Title(i).EndFrame = Title(i).EndFrame + Change
    Next i
    Screen.MousePointer = 0
End Sub

'--- Changes the FPS of all titles from iFrom to iTo ---
Sub ChangeFPS(ByVal iFrom As Single, ByVal iTo As Single)
    Dim i As Long
    Dim dMult As Double
    dMult = iTo / iFrom
    For i = 1 To LineCnt
        Title(i).StartFrame = Round(Title(i).StartFrame * dMult, 0)
        Title(i).EndFrame = Round(Title(i).EndFrame * dMult)
    Next i
End Sub

'--- Adjusts starting and ending position of a title (adds the Change) ---
Sub ChangeTitle(ByVal lIndex As Long, bIsStart As Boolean, ByVal Change As Double)
    If bIsStart Then
        Title(lIndex).StartFrame = Round(Title(lIndex).StartFrame + Change)
    Else
        Title(lIndex).EndFrame = Round(Title(lIndex).EndFrame + Change)
    End If
End Sub

'--- Deletes a title from the array ---
Sub DeleteTitle(lIndex As Long)
    Dim i As Long
    If LineCnt < 1 Then Exit Sub
    For i = lIndex To LineCnt - 1
        Title(i) = Title(i + 1)
    Next i
    LineCnt = LineCnt - 1
    ReDim Preserve Title(1 To LineCnt)
End Sub

'--- inserts an empty title into the array at lIndex position ---
Sub InsertTitle(ByRef lIndex As Long)
    Dim i As Long
    LineCnt = LineCnt + 1
    ReDim Preserve Title(1 To LineCnt)
    If lIndex < LineCnt Then
        For i = LineCnt To lIndex + 1 Step -1
            Title(i) = Title(i - 1)
        Next i
    Else
        lIndex = LineCnt
    End If
End Sub

'--- Append titles to a file ---
Sub AppendTitles(ByVal sFileName As String)
    Dim tOld() As Subtitles
    Dim sNewFile As String
    GetTitles sFileName, tOld()
    sNewFile = Left(sFileName, Len(sFileName) - 4) + "-NEW" + Right(sFileName, 4)
    SaveTitles sNewFile, tOld()
    SaveTitles sNewFile, Title(), True
    Erase tOld
End Sub

'--- Reads settings ---
Sub GetSettings()
    lUpdateTime = Val(GetSetting(cst_AppName, cst_Settings, ckey_UpdateTime, "500"))
    lMeasureTime = Val(GetSetting(cst_AppName, cst_Settings, ckey_MeasureTime, "20000"))
    bAutoMeasure = GetSetting(cst_AppName, cst_Settings, ckey_AutoMeasure, "False") = "True"
End Sub

'--- Saves settings ---
Sub SaveSettings()
    SaveSetting cst_AppName, cst_Settings, ckey_UpdateTime, lUpdateTime
    SaveSetting cst_AppName, cst_Settings, ckey_MeasureTime, lMeasureTime
    SaveSetting cst_AppName, cst_Settings, ckey_AutoMeasure, bAutoMeasure
End Sub

Sub Main()
    GetSettings
    frmSub.Show
End Sub

'--- Generates a MicroDVD Player .ini file and an Autorun.inf for a CD ---
Sub GenerateMDVD(ByVal sSubFile As String, ByVal sAviFile As String, ByVal sDestDir As String, Optional sMovieTitle = "", Optional lCDNo = 1)
    Dim sAutorun As String, sMDVD As String
    sDestDir = BF.ToPath(sDestDir)
    sAutorun = sDestDir + "autorun.inf"
    BF.WriteProfileString sAutorun, "Autorun", "Open", "MDVD.MVD"
    sMDVD = sDestDir + "mdvd.ini"
    Open sMDVD For Output As #1
    Print #1, "[Micro DVD Ini File]"
    Close #1
    BF.WriteProfileString sMDVD, "MAIN", "TITLE", sMovieTitle
    BF.WriteProfileString sMDVD, "MAIN", "ID", CStr(lCDNo)
    BF.WriteProfileString sMDVD, "MAIN", "CDNumber", "1"
    BF.WriteProfileString sMDVD, "MAIN", "Delay", "1"
    BF.WriteProfileString sMDVD, "MOVIE", "Directory", "."
    BF.WriteProfileString sMDVD, "MOVIE", "AVIName", BF.GetFileNameFromDir(sAviFile)
    BF.WriteProfileString sMDVD, "SUBTITLES", "Directory", "."
    BF.WriteProfileString sMDVD, "SUBTITLES", "DialogString", "-"
    BF.WriteProfileString sMDVD, "SUBTITLES", "1", "BUL Bulgarian" 'type your language here (3 letters short name + space + long name)
    BF.WriteProfileString sMDVD, "SUBTITLES", "File", BF.GetFileNameFromDir(sSubFile)
    BF.WriteProfileString sMDVD, "SUBTITLES", "Offset", "0"
    BF.WriteProfileString sMDVD, "SUBTITLES", "Multiplier", "100000"
    If Dir(sDestDir + "mdvd.mvd") <> "" Then Kill sDestDir + "mdvd.mvd"
    FileCopy sDestDir + "mdvd.ini", sDestDir + "mdvd.mvd"
End Sub

'--- Converts frames to time code (.SRT format) ---
Function Convert2Time(lFrame As Long) As String
    Dim dTime As Double, dMilliSecs As Double, dTimeInt As Long
    dTime = (lFrame / gFPS)
    dTimeInt = Int(dTime)
    dMilliSecs = Int(1000 * (dTime - dTimeInt))
    Convert2Time = Format(dTimeInt / 86400, "hh:mm:ss") + "," + Format(dMilliSecs, "000")
End Function

'--- Convert time code to frames (.SRT format) ---
Function ConvertFromTime(sTime As Variant) As Long
    Dim dTime As Double, dMilliSecs As Long
    Dim t
    dMilliSecs = Val(Right(sTime, 3))
    t = Split(sTime, ":")
    dTime = (Val(t(0)) * 3600 + Val(t(1)) * 60 + Val(t(2)) + dMilliSecs / 1000)
    ConvertFromTime = Round(gFPS * dTime)
End Function
