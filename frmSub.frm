VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subtitler"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12630
   Icon            =   "frmSub.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   12630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmSub 
      Height          =   8070
      Left            =   4185
      TabIndex        =   24
      Top             =   45
      Width           =   8430
      Begin VB.TextBox txFile 
         Height          =   285
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   " "
         Top             =   225
         Width           =   6690
      End
      Begin VB.CommandButton btSubBrowse 
         Caption         =   "..."
         Height          =   270
         Left            =   7740
         TabIndex        =   26
         ToolTipText     =   "Load subtitle file"
         Top             =   225
         Width           =   540
      End
      Begin MSComDlg.CommonDialog CDL 
         Left            =   6030
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "sub"
         Filter          =   "MicroDVD Subtitles|*.sub|SubRip Subtitles|*.srt"
      End
      Begin MSFlexGridLib.MSFlexGrid GRID 
         Height          =   7485
         Left            =   90
         TabIndex        =   25
         Top             =   540
         Visible         =   0   'False
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   13203
         _Version        =   393216
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         Caption         =   "Subtitle file:"
         Height          =   225
         Left            =   135
         TabIndex        =   29
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SUBTITLES NOT LOADED."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2025
         TabIndex        =   28
         Top             =   3555
         Width           =   4830
      End
   End
   Begin VB.Frame fmMovie 
      Caption         =   "Movie file:"
      Height          =   5325
      Left            =   0
      TabIndex        =   19
      Top             =   45
      Width           =   4155
      Begin VB.CommandButton btMovieBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   3600
         TabIndex        =   23
         ToolTipText     =   "Load movie file"
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox txMovie 
         Height          =   285
         Left            =   45
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   3525
      End
      Begin VB.OptionButton opSelect 
         Caption         =   "Out"
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1710
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3825
         Width           =   465
      End
      Begin VB.OptionButton opSelect 
         Caption         =   "In"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3825
         Width           =   465
      End
      Begin VB.Label lbSub 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   960
         Left            =   90
         TabIndex        =   20
         Top             =   4275
         Width           =   3930
      End
      Begin MediaPlayerCtl.MediaPlayer MMPlayer 
         Height          =   3165
         Left            =   90
         TabIndex        =   21
         Top             =   630
         Width           =   3930
         AudioStream     =   -1
         AutoSize        =   0   'False
         AutoStart       =   -1  'True
         AnimationAtStart=   0   'False
         AllowScan       =   -1  'True
         AllowChangeDisplaySize=   -1  'True
         AutoRewind      =   0   'False
         Balance         =   0
         BaseURL         =   ""
         BufferingTime   =   3
         CaptioningID    =   ""
         ClickToPlay     =   -1  'True
         CursorType      =   0
         CurrentPosition =   -1
         CurrentMarker   =   0
         DefaultFrame    =   ""
         DisplayBackColor=   0
         DisplayForeColor=   16777215
         DisplayMode     =   1
         DisplaySize     =   4
         Enabled         =   -1  'True
         EnableContextMenu=   -1  'True
         EnablePositionControls=   -1  'True
         EnableFullScreenControls=   0   'False
         EnableTracker   =   -1  'True
         Filename        =   ""
         InvokeURLs      =   0   'False
         Language        =   -1
         Mute            =   0   'False
         PlayCount       =   1
         PreviewMode     =   -1  'True
         Rate            =   1
         SAMILang        =   ""
         SAMIStyle       =   ""
         SAMIFileName    =   ""
         SelectionStart  =   -1
         SelectionEnd    =   -1
         SendOpenStateChangeEvents=   -1  'True
         SendWarningEvents=   -1  'True
         SendErrorEvents =   -1  'True
         SendKeyboardEvents=   0   'False
         SendMouseClickEvents=   0   'False
         SendMouseMoveEvents=   0   'False
         SendPlayStateChangeEvents=   -1  'True
         ShowCaptioning  =   0   'False
         ShowControls    =   -1  'True
         ShowAudioControls=   -1  'True
         ShowDisplay     =   0   'False
         ShowGotoBar     =   0   'False
         ShowPositionControls=   -1  'True
         ShowStatusBar   =   -1  'True
         ShowTracker     =   -1  'True
         TransparentAtStart=   0   'False
         VideoBorderWidth=   0
         VideoBorderColor=   0
         VideoBorder3D   =   0   'False
         Volume          =   0
         WindowlessVideo =   0   'False
      End
      Begin VB.Label lbDuration 
         Alignment       =   2  'Center
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   3375
         TabIndex        =   35
         Top             =   3915
         Width           =   735
      End
      Begin VB.Label lbOut 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2205
         TabIndex        =   34
         Top             =   3870
         Width           =   1140
      End
      Begin VB.Label lbIn 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   585
         TabIndex        =   33
         Top             =   3870
         Width           =   1050
      End
   End
   Begin VB.Frame fmMark 
      Height          =   2760
      Left            =   0
      TabIndex        =   1
      Top             =   5355
      Width           =   4155
      Begin VB.OptionButton opMark 
         Caption         =   "1"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "2"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   405
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "3"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "4"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1035
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "5"
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "6"
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1665
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1530
         Width           =   330
      End
      Begin VB.OptionButton opMark 
         Caption         =   "7"
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1530
         Width           =   330
      End
      Begin VB.CommandButton btMark 
         Caption         =   "Set"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2610
         TabIndex        =   5
         Top             =   1530
         Width           =   690
      End
      Begin VB.CommandButton btMark 
         Caption         =   "Go to"
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3330
         TabIndex        =   4
         Top             =   1530
         Width           =   690
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2430
         Top             =   2250
      End
      Begin VB.ComboBox cbFPS 
         Height          =   315
         ItemData        =   "frmSub.frx":08CA
         Left            =   1575
         List            =   "frmSub.frx":08E3
         TabIndex        =   3
         Text            =   "25"
         Top             =   180
         Width           =   825
      End
      Begin VB.CommandButton btMeasure 
         Caption         =   "Measure"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2790
         TabIndex        =   2
         Top             =   180
         Width           =   1230
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   3285
         Top             =   2250
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   3690
         Top             =   2250
      End
      Begin MSComctlLib.Slider slSpeed 
         Height          =   420
         Left            =   90
         TabIndex        =   13
         Top             =   1935
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   741
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         Max             =   4
         SelStart        =   2
         TickStyle       =   1
         Value           =   2
      End
      Begin MSComDlg.CommonDialog CDL2 
         Left            =   3375
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Load a movie file"
         Filter          =   "Windows Media Files|*.avi;*.asf|MPEG files|*.mpg;*.mpe;*.mpeg|All files|*.*"
      End
      Begin VB.Label lbFrame 
         BackColor       =   &H00000000&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   30
         Top             =   540
         Width           =   3930
      End
      Begin VB.Label Label5 
         Caption         =   "Position markers:"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   1305
         Width           =   3255
      End
      Begin VB.Label lbInfo 
         AutoSize        =   -1  'True
         Caption         =   "Playback speed: 100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   17
         Top             =   2385
         Width           =   1890
      End
      Begin VB.Label lbFrame 
         BackColor       =   &H00000000&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   900
         Width           =   3930
      End
      Begin VB.Label Label2 
         Caption         =   "Movie frame rate is:"
         Height          =   240
         Left            =   135
         TabIndex        =   15
         Top             =   225
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "FPS"
         Height          =   240
         Left            =   2430
         TabIndex        =   14
         Top             =   225
         Width           =   420
      End
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8175
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7250
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7250
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "File"
      Begin VB.Menu mnNew 
         Caption         =   "New subtitle file"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnOpen 
         Caption         =   "Open video file..."
         Index           =   0
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnOpen 
         Caption         =   "Open subtitle file..."
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnMerge 
         Caption         =   "Append subtitle file..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSave 
         Caption         =   "Save subtitles as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnAppend 
         Caption         =   "Append subtitles to..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnSubEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnInsert 
         Caption         =   "Insert row before"
         Index           =   0
         Shortcut        =   ^I
      End
      Begin VB.Menu mnInsert 
         Caption         =   "Add new row"
         Index           =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnDelete 
         Caption         =   "Remove row"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnCopy 
         Caption         =   "Copy title"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnPaste 
         Caption         =   "Paste title"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnPasteFrames 
         Caption         =   "Paste frames"
         Enabled         =   0   'False
         Shortcut        =   ^W
      End
      Begin VB.Menu mnSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnGoto 
         Caption         =   "Set movie at current position"
         Enabled         =   0   'False
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnFind 
         Caption         =   "Find & Replace..."
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnFind 
         Caption         =   "Find next"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnSplit 
         Caption         =   "Split subtitles..."
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnAutoInsert 
         Caption         =   "Auto-insert mode"
         Checked         =   -1  'True
         Shortcut        =   {F11}
      End
   End
   Begin VB.Menu mnTools 
      Caption         =   "Tools"
      Begin VB.Menu mnMDVD 
         Caption         =   "Generate MDVD.INI"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnCinema 
         Caption         =   "Cinema mode"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnChangeSettings 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------- MAIN WINDOW -----------

Option Explicit

Dim cPos As Long, cTime As String
Dim iState As Integer
Dim cMarker As Integer
Dim bMeasureFPS  As Boolean
Dim LastTitle As Long
Dim Probe() As Single
Dim ProbeCnt As Long

'--- sets or retrieves a bookmark position ---
Private Sub btMark_Click(Index As Integer)
    If cMarker = 0 Then Exit Sub
    Select Case Index
    Case 0
        Mark(cMarker) = MMPlayer.CurrentPosition
        opMark(cMarker - 1).ToolTipText = "Frame " + CStr(Mark(cMarker))
    Case 1
        MMPlayer.CurrentPosition = Mark(cMarker)
    End Select
End Sub

'---  Starts measuring the movie's frame rate (FPS)
'     if you know better way to do it , please email me!!
Private Sub btMeasure_Click()
    On Error GoTo 100
    Screen.MousePointer = 11
    sbStatus.Panels(1).Text = "Measuring FPS... Please, wait"
    ProbeCnt = 0
    gFPS = 0
    'play the movie if it is not already playing:
    If iState <> 2 Then
        MMPlayer.Play
    End If
    'waits for a couple of seconds after movie is started:
    Timer2.Enabled = True
10
    Exit Sub
100
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

'--- Shows the Open dialog and loads movie into the Media Player window ---
Private Sub btMovieBrowse_Click()
    Dim i As Integer
    On Error GoTo 100
    With CDL2
        .ShowOpen
        sAviFile = .FileName  'remember movie's name
        txMovie.Text = BF.GetFileNameFromDir(sAviFile)
    End With
    'loads the movie into the player and tries to find subtitle file with matching name:
    LoadMovieFile
100
    Exit Sub
200
    MsgBox Err.Description, vbCritical
    Resume 100
End Sub

'--- Shows the open dialog and loads the subtitle file ---
Private Sub btSubBrowse_Click()
    'On Error GoTo 100
    CDL.ShowOpen
    'On Error Resume Next
    'processes the subtitle file and loads it into the Title array
    LoadSubTitleFile CDL.FileName
100
End Sub

'--- Shows the 'Save as' dialog and writes subtitles to the disk (2 formats available) ---
Private Sub SaveSubtitles(iAppend As Integer)
    Dim sOld As String, sNewFile As String
    On Error Resume Next
    With CDL
        .FileName = Left(sSubFileName, Len(sSubFileName) - 4)
        On Error GoTo 100
        .ShowSave
        On Error GoTo 200
        'unsets the 'read-only' attribute (it is often set when the file has been copied from a CD)
        If Dir(.FileName) <> "" Then SetAttr .FileName, vbNormal
        '
        If iAppend = 0 Then
            SaveTitles .FileName, Title() 'saves to a new file
        Else
            AppendTitles .FileName        'appends to an existing file
        End If
        sSubFileName = .FileName
        txFile.Text = sSubFileName
        MsgBox sSubFileName + " succesfuly written to the disk.", vbInformation
    End With
    bChanged = False
100
    Exit Sub
200
    MsgBox Err.Description, vbCritical
    Resume 100
End Sub



Private Sub MMPlayer_OpenStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If NewState = 6 And bAutoMeasure Then btMeasure_Click
End Sub

Private Sub mnAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnAutoInsert_Click()
    mnAutoInsert.Checked = Not mnAutoInsert.Checked
End Sub

'--- Switches to the 'Cinema' mode (allows to use the program as a simple player) ---
Private Sub mnCinema_Click()
    Static bOn As Boolean
    bOn = Not bOn
    mnCinema.Checked = bOn
    fmSub.Visible = Not bOn
    fmMark.Visible = Not bOn
    opSelect(0).Visible = Not bOn
    opSelect(1).Visible = Not bOn
    'resizes controls on the form
    If bOn Then
        fmMovie.Move 0, 45, Me.ScaleWidth - 50, Me.ScaleHeight - 300
        lbSub.Move 90, fmMovie.Height - 1250, fmMovie.Width - 200, 1200
        MMPlayer.Width = fmMovie.Width - 200
        MMPlayer.Height = fmMovie.Height - lbSub.Height - txMovie.Height - 400
        lbSub.Font.Size = 16
    Else
        fmMovie.Move 0, 45, 4155, 5325
        MMPlayer.Width = 3930
        MMPlayer.Height = 3165
        lbSub.Move 90, 4275, 3930, 960
        lbSub.Font.Size = 10
    End If
End Sub

'--- Copies an entire title to the Clipboard ---
Private Sub mnCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Title(GRID.Row).Text
End Sub

'--- Exits the program ---
Private Sub mnExit_Click()
    MMPlayer.Stop
    Unload Me
End Sub

'--- Manualy sets the FPS rate ---
Private Sub cbFPS_Change()
    gFPS = Val(cbFPS.Text)
End Sub

Private Sub cbFPS_Click()
     gFPS = Val(cbFPS.Text)
End Sub

'--- Some keyboard shortcuts ---
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lIndex As Long
    Select Case KeyCode
    Case vbKeyHome 'goes to the top
        GRID.TopRow = 1
        GRID.Row = 1
        GoFrame 1
    Case vbKeyEnd 'goes to the bottom
        GRID.TopRow = LineCnt
        GRID.Row = LineCnt
        GoFrame LineCnt
    
    Case vbKeySpace 'pauses/plays the movie
        If iState <> 2 Then MMPlayer.Play Else MMPlayer.Pause
    Case vbKeyAdd   'inserts a new line
        mnInsert_Click (Abs(Shift = 0))
    Case vbKeyMultiply 'sets movie position at the current title
        GoFrame lSelStart
    Case vbKeySubtract
        lIndex = GRID.Row
        frmEdit.SetTitle lIndex
        frmEdit.Show 1
    Case vbKeyF5 'marks in
        opSelect_Click (0)
    Case vbKeyF6 'marks out
        opSelect_Click (1)
    Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7  'set/retrieve position  markers
        lIndex = KeyCode - vbKey1
        opMark(lIndex).Value = True
        cMarker = lIndex
        If Shift = 1 Then 'press SHIFT+number to set position
            btMark_Click (0)
        ElseIf Shift = 2 Then 'press CTRL+number to retrieve position
            btMark_Click (1)
        End If
    End Select
    KeyCode = 0
End Sub

'--- Initializes the program and processes the command line ---
Private Sub Form_Load()
    Dim sExt As String, sCommand As String
    gFPS = 25
    InitSettings
    LastTitle = 1
    mnNew_Click
    'if a command line exists, tries to load a file
    If Command <> "" Then
        'remove "" marks
        sCommand = Replace(Command, Chr(34), "")
        'get the extension
        sExt = LCase(BF.rightPart(sCommand, ".", True))
        Select Case sExt
        Case "avi" 'load a movie
            sAviFile = sCommand
            LoadMovieFile
        Case "sub", "srt" 'load a subtitle file
            LoadSubTitleFile sCommand
        End Select
    End If
End Sub

'--- Clean-up on exit ---
Private Sub Form_Unload(Cancel As Integer)
    Set BF = Nothing
    End
End Sub

'--- Enables/disables menus when title is selected---
Private Sub GRID_Click()
    If GRID.Row > 0 Then
        lSelStart = GRID.Row
        mnSplit.Enabled = GRID.Row > 1 And GRID.Row < LineCnt
        mnDelete.Enabled = True
    Else
        mnSplit.Enabled = False
        mnDelete.Enabled = False
    End If
End Sub

'--- Pops up a dialog according to the column you have clicked on ---
Private Sub GRID_DblClick()
    Dim MouseIndex As Integer
    If GRID.Rows < 2 Then Exit Sub
    MouseIndex = GRID.MouseCol
    lSelStart = GRID.MouseRow
    'first 2 columns show the frame calculator
    If MouseIndex = 1 Or MouseIndex = 2 Then
        bTitleStart = MouseIndex = 1
        If bTitleStart Then
            frmAction.Text1 = CStr(Title(lSelStart).StartFrame)
        Else
            frmAction.Text1 = CStr(Title(lSelStart).EndFrame)
        End If
        frmAction.Show 1
        'the third column shows the title editor
    ElseIf MouseIndex = 3 Then
        frmEdit.SetTitle lSelStart
        frmEdit.Show 1
    End If
End Sub

Private Sub GRID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GoFrame GRID.Row
    End If
End Sub

'--- Shows the pop-up menu on right-click ---
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        GRID.Row = GRID.MouseRow
        PopupMenu mnSubEdit
    End If
End Sub

Private Sub lbSub_Click()
    On Error Resume Next
    If GRID.Rows > 1 And lSelStart > 0 Then
        frmEdit.SetTitle lSelStart
        frmEdit.Show 1
    End If
End Sub

'--- Updates the status bar and the counters when the movie state has changed ---
Private Sub MMPlayer_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    'remember the state:
    iState = NewState
    Timer1.Enabled = iState = 2
    lbSub.Caption = ""
    LastTitle = 1
    ShowTitle
    cPos = MMPlayer.CurrentPosition
    lbFrame(0).Caption = "Frame: " + CStr(cPos)
    sbStatus.Panels(1).Text = Choose(iState + 1, "Stopped.", "Paused.", "Playing...")
End Sub

'--- Updates the frame counter and shows the corresponding title ---
Private Sub MMPlayer_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
    ShowTitle
    cPos = MMPlayer.CurrentPosition
    lbFrame(0).Caption = "Frame: " + CStr(cPos)
End Sub

'--- Appends the titles to an existing file ---
Private Sub mnAppend_Click()
    SaveSubtitles 1
End Sub

'--- Shows the settings dialog ---
Private Sub mnChangeSettings_Click()
    frmSettings.Show 1
End Sub

'--- Deletes a title ---
Private Sub mnDelete_Click()
    If MsgBox("Do you really wish to delete row No. " + CStr(GRID.Row) + "?", vbExclamation + vbYesNo) = vbYes Then
        DeleteTitle GRID.Row
        RefreshTitles
    End If
End Sub

'--- Shows the Find & Replace dialog ---
Private Sub mnFind_Click(Index As Integer)
    frmFind.Show
End Sub

'--- Sets player to the starting frame of the current title ---
Private Sub mnGoto_Click()
    GoFrame GRID.MouseRow
End Sub

Private Sub mnHelpContents_Click()
    If Not (BF.ShellOpen(Me.hwnd, BF.ToPath(App.Path) + "help.htm")) Then
        MsgBox "Help file is not installed.", vbCritical
    End If
End Sub

'--- Inserts or appends a title ---
Private Sub mnInsert_Click(Index As Integer)
    lSelStart = GRID.Row
    Select Case Index
    Case 0 'Insert
        
    Case 1 'Add to end
        lSelStart = LineCnt + 1
    End Select
    frmEdit.InsertAt lSelStart
End Sub

'--- Generates a MicroDVD player .ini file ---
Private Sub mnMDVD_Click()
    Dim sDir As String
    sDir = BF.BrowseFolder(Me.hwnd, "Select destination folder for your movie files:")
    If sDir <> "" Then
        GenerateMDVD sSubFileName, sAviFile, sDir
    End If
    If MsgBox("The Mdvd.ini file has been generated into " + sDir + vbCrLf + "Do you wish to view it now?", vbYesNo) = vbYes Then
        Shell "Notepad.exe " + BF.ToPath(sDir) + "mdvd.ini", vbNormalFocus
    End If
End Sub

'--- Loads subtitles from a file and merges it to the currently open file ---
Private Sub mnMerge_Click()
    Dim tTmp() As Subtitles, NewCnt As Long
    Dim OldCnt As Long, FirstFrame As Long, Change As Long
    Dim i As Long
    'show the Open dialog
    On Error GoTo 100
    CDL.ShowOpen
    On Error GoTo 200
    'get the number of the newly loaded titles
    NewCnt = GetTitles(CDL.FileName, tTmp())
    If NewCnt > 0 Then
        'get the merging position
        FirstFrame = tTmp(1).StartFrame
        'get the frame offset
        Change = Title(LineCnt).EndFrame + 2
        OldCnt = LineCnt + 1
        'add the number of titles to the title counter
        LineCnt = LineCnt + NewCnt
        'redimension the Title array
        ReDim Preserve Title(1 To LineCnt)
        'write to the title array
        For i = OldCnt To LineCnt
            Title(i) = tTmp(i - OldCnt + 1)
            'auto-calculate the frames offset
            Title(i).StartFrame = Title(i).StartFrame + Change
            Title(i).EndFrame = Title(i).EndFrame + Change
        Next i
    End If
    'show changes into the grid:
    RefreshTitles
    GRID.TopRow = OldCnt
    GRID.Row = OldCnt
    'mark the merging point in yellow:
    For i = 0 To GRID.Cols - 1
        GRID.Col = i
        GRID.CellBackColor = vbYellow
    Next i
    MsgBox "File " + CDL.FileName + " has been appended at position " + CStr(OldCnt), vbInformation
100
    Exit Sub
200
    MsgBox Err.Description, vbCritical
    Resume 100
End Sub

'--- Prepares a new subtitle file ---
Private Sub mnNew_Click()
    'ReDim Title()
    LineCnt = 0
    RefreshTitles
    txFile.Text = "<not saved>"
    GRID.Visible = True
End Sub

'--- Opens a movie / subtitle file ---
Private Sub mnOpen_Click(Index As Integer)

    Select Case Index
    Case 0
        btMovieBrowse_Click
    Case 1
        btSubBrowse_Click
    End Select
End Sub

'--- Pastes a previously copied title from Clipboard ---
Private Sub mnPaste_Click()
    Title(GRID.Row).Text = Clipboard.GetText
    GRID.TextMatrix(GRID.Row, 3) = Title(GRID.Row).Text
End Sub

'--- Pastes the frame interval measured by In and Out buttons to the current title ---
Private Sub mnPasteFrames_Click()
    If lSelStart > 0 Then
        If IsNumeric(lbOut.Caption) Then
            Title(lSelStart).StartFrame = Val(lbIn.Caption)
            Title(lSelStart).EndFrame = Val(lbOut.Caption)
        End If
    End If
End Sub

'--- Saves the subtitles ---
Private Sub mnSave_Click()
    SaveSubtitles 0
End Sub

'--- Shows the Split subtitles dialog ---
Private Sub mnSplit_Click()
    If cPos = 0 Then
        frmSplit.SetPosition Title(lSelStart).EndFrame, lSelStart
    Else
        frmSplit.SetPosition cPos, GetSplitPosition
    End If
    frmSplit.Show 1
End Sub

'--- Assigns /Retrieves current movie position to/from a position marker ---
Private Sub opMark_Click(Index As Integer)
    On Error Resume Next
    btMark(0).Enabled = True
    btMark(1).Enabled = True
    cMarker = Index + 1
    'remember position
    If Mark(cMarker) = 0 Then
        MMPlayer.Pause
        Mark(cMarker) = MMPlayer.CurrentPosition
        opMark(Index).Caption = CStr(cMarker) + "*"
    Else
    'retrieve position
        MMPlayer.Pause
        MMPlayer.CurrentPosition = Mark(cMarker)
    End If
End Sub

'--- Gets the In or Out position of the subtitle ---
Private Sub opSelect_Click(Index As Integer)
    Select Case Index
    Case 0 'mark in (starting frame)
        If Val(lbOut.Caption) > MMPlayer.CurrentPosition Then MMPlayer.CurrentPosition = Val(lbOut.Caption) + 1
        lbIn.Caption = MMPlayer.CurrentPosition
        lbOut.Caption = "...."
        lbDuration.Caption = "..."
        mnPaste.Enabled = False
        MMPlayer.Play
    Case 1 'mark out (ending frame)
        lbOut.Caption = MMPlayer.CurrentPosition
        MMPlayer.Pause
        lbDuration = Format((Val(lbOut.Caption) - Val(lbIn.Caption)) / gFPS, "0.0") + " s"
        'if auto-insertion option is set, auto-insert a new title
        If mnAutoInsert.Checked Then
            lSelStart = lSelStart + 1
            frmEdit.InsertAt lSelStart, Val(lbIn.Caption), Val(lbOut.Caption)
        End If
        mnPaste.Enabled = True
    End Select
End Sub

Private Sub opSelect_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If opSelect(0).Value Then
        lbIn.Caption = "...."
        opSelect(0).Value = False
        Exit Sub
    End If
End Sub

'--- Changes the playback speed of the movie ---
Private Sub slSpeed_Change()
    MMPlayer.Rate = slSpeed.Value / 2
    lbInfo(0).Caption = "Playback speed: " + Format(MMPlayer.Rate, "0%")
End Sub

'--- Fills the grid with the information contained in the Title array ---
Sub RefreshTitles()
    Dim i As Long
    Dim Wdth, Head
    Dim lTopRow As Long
    On Error Resume Next
    Screen.MousePointer = 11
    'get column headings and widths
    Wdth = Array(500, 1000, 1000, 5300)
    Head = Array("No", "Start", "End", "Subtitle")
    
    'remember the top row, so it can be set again later
    If GRID.Rows > 1 Then lTopRow = GRID.TopRow
    If Err Then lTopRow = 1: Err.Clear
    'stop refreshing the grid
    GRID.Redraw = False
    'clear the grid
    GRID.Clear
    GRID.Cols = 4
    GRID.Rows = 1
    'put on the column headings and resize the columns:
    For i = 0 To GRID.Cols - 1
        GRID.ColWidth(i) = Wdth(i)
        GRID.TextMatrix(0, i) = Head(i)
        GRID.ColAlignment(i) = 1
    Next i
    'start filling in the titles:
    For i = 1 To LineCnt
        'append a new row
        GRID.Rows = i + 1
        'write the index
        GRID.TextMatrix(i, 0) = CStr(i)
        'write the starting and ending frames
        GRID.TextMatrix(i, 1) = CStr(Title(i).StartFrame)
        GRID.TextMatrix(i, 2) = CStr(Title(i).EndFrame)
        'write the subtitle text
        GRID.TextMatrix(i, 3) = Title(i).Text
        'if titles overlap, mark them as red
        If i > 1 Then
            If Title(i).StartFrame < Title(i - 1).EndFrame Then
                GRID.Col = 1
                GRID.Row = i
                GRID.CellBackColor = vbRed
                GRID.Col = 2
                GRID.Row = i - 1
                GRID.CellBackColor = vbRed
            End If
        End If
    Next i
    'reset the top row
    GRID.TopRow = lTopRow
    GRID.Row = lSelStart
10
    'start redrawing the grid
    GRID.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
100
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

'--- Loads the subtitles from a file to memory ---
Sub LoadTitles(sFileName As String)
    LineCnt = GetTitles(sFileName, Title())
    'and shows them
    RefreshTitles
End Sub

'--- Shows the title matching the current player's position at the subtitle window ---
Sub ShowTitle()
    Dim i As Long
    On Error Resume Next
    For i = LastTitle To LineCnt
        If Title(i).StartFrame <= cPos And Title(i).EndFrame >= cPos Then
            If Not GRID.RowIsVisible(GRID.Rows - 1) Then
                GRID.TopRow = i
            End If
            GRID.Row = i
            lSelStart = i
            lbSub.Caption = Replace(Title(i).Text, "|", vbCrLf)
            LastTitle = i
            Exit Sub
        End If
    Next i
    lbSub.Caption = ""
End Sub

'--- Sets the player's position at the starting frame of the current title ---
Sub GoFrame(ByVal lIndex As Long)
    lSelStart = lIndex
    LastTitle = lIndex
    lbSub.Caption = ""
    MMPlayer.CurrentPosition = Title(lSelStart).StartFrame
    MMPlayer.Stop
End Sub


'--- Starts the FPS measuring in a couple of seconds the avoid the initial loading delay ---
Private Sub Timer2_Timer()
    bMeasureFPS = True
    Timer3.Enabled = True
End Sub

'--- During a playback, takes probes of the elapsed frames, which can be used for FPS measuring ---
Private Sub Timer1_Timer()
    Static OldPos As Long
    Dim dRate As Double
    cPos = MMPlayer.CurrentPosition
    If bMeasureFPS Then
        'adjust the time interval for the playback rate
        dRate = MMPlayer.Rate / (Timer1.Interval / 1000)
        gFPS = Round((cPos - OldPos) * dRate, 0)
        'remember the probe
        ProbeCnt = ProbeCnt + 1
        ReDim Preserve Probe(ProbeCnt)
        Probe(ProbeCnt) = gFPS
        If OldPos <> cPos Then lbFrame(1).Caption = "Frame rate: " + CStr(gFPS)
    Else
        ShowTitle
    End If
    OldPos = cPos
End Sub

'--- Estimates the FPS after the pre-defined time interval has elapsed  ---
Private Sub Timer3_Timer()
    Dim i As Integer, AvgFPS As Single
    bMeasureFPS = False
    'stop the timers:
    Timer2.Enabled = False
    Timer3.Enabled = False
    'calculate the average frames per second based on the number of probes taken
    For i = 1 To ProbeCnt
        AvgFPS = AvgFPS + Probe(i)
    Next i
    gFPS = Round(AvgFPS / ProbeCnt, 3)
    Erase Probe
    'stop the movie
    MMPlayer.Stop
    Screen.MousePointer = 0
    'show the result
    lbFrame(1).Caption = "Frame rate: " + CStr(gFPS)
    sbStatus.Panels(1).Text = "Movie FPS rate measured. (" + CStr(gFPS) + " FPS)"
    cbFPS.Text = CStr(gFPS)
End Sub

'--- Assign time intervals to the timers ---
Sub InitSettings()
    Timer1.Interval = lUpdateTime
    Timer3.Interval = lMeasureTime
End Sub

Function GetSplitPosition() As Long
    Dim i As Long
    cPos = MMPlayer.CurrentPosition
    For i = 1 To LineCnt - 1
        If Title(i).StartFrame <= cPos And Title(i + 1).StartFrame > cPos Then
            GetSplitPosition = i
            Exit Function
        End If
    Next i
    If (cPos >= Title(LineCnt).StartFrame) Then GetSplitPosition = LineCnt
End Function

'--- Processes the subtitle loading ---
Private Sub LoadSubTitleFile(ByVal sFileName As String)
    AskToSave
    sSubFileName = sFileName
    LoadTitles sSubFileName
    'attempt to load a movie with the same name
    If Dir(Left(sSubFileName, Len(sSubFileName) - 3) + "avi") <> "" And sAviFile = "" Then
        sAviFile = Left(sSubFileName, Len(sSubFileName) - 3) + "avi"
        LoadMovieFile
    End If
    'update the status bar
    sbStatus.Panels(2).Text = "File: " + BF.GetFileNameFromDir(sSubFileName) + " (" + CStr(LineCnt) + " lines)"
    'Enable/disable controls and menus
    mnGoto.Enabled = sAviFile <> ""
    mnMDVD.Enabled = sAviFile <> ""
    mnFind(0).Enabled = True
    mnFind(1).Enabled = True
    GRID.Visible = True
    lSelStart = 1
    txFile.Text = sSubFileName
End Sub

'--- Processes the movie loading ---
Private Sub LoadMovieFile()
    Dim i As Integer
    On Error GoTo 200
    'reset the player
    MMPlayer.AutoStart = False
    MMPlayer.ShowGotoBar = True
    MMPlayer.Open sAviFile
    'set the most common FPS
    gFPS = 25
    'enable/disable controls and menus
    Timer1.Enabled = False
    btMeasure.Enabled = True
    opSelect(0).Enabled = True
    opSelect(1).Enabled = True
    btMark(0).Enabled = True
    btMark(1).Enabled = True
    slSpeed.Enabled = True
    mnMDVD.Enabled = sSubFileName <> ""
    'attempt to load a subtitle file with the same name
    If Dir(Left(sAviFile, Len(sAviFile) - 3) + "sub") <> "" Then
        sSubFileName = Left(sAviFile, Len(sAviFile) - 3) + "sub"
        LoadSubTitleFile sSubFileName
    End If
    'reset position markers
    For i = 0 To opMark.Count - 1
        opMark(i).Enabled = True
    Next i
    mnGoto.Enabled = sSubFileName <> ""
10
    Exit Sub
200
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub

'--- Warns the user that the subtitles have not been saved ---
Sub AskToSave()
    If bChanged Then
        If MsgBox("The subtitles have not been saved. Do you wish to save them now?", vbYesNo + vbExclamation) = vbYes Then
            SaveSubtitles 0
        End If
    End If
End Sub
