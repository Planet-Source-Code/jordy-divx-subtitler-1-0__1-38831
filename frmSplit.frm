VERSION 5.00
Begin VB.Form frmSplit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Split subtitles"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4380
   Icon            =   "frmSplit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2835
      TabIndex        =   6
      Top             =   1845
      Width           =   1500
   End
   Begin VB.CommandButton btSave 
      Caption         =   "Save files"
      Default         =   -1  'True
      Height          =   330
      Left            =   45
      TabIndex        =   5
      Top             =   1845
      Width           =   1500
   End
   Begin VB.TextBox txPos 
      Height          =   285
      Index           =   1
      Left            =   1755
      TabIndex        =   4
      Text            =   "0000"
      Top             =   1350
      Width           =   1230
   End
   Begin VB.TextBox txPos 
      Height          =   285
      Index           =   0
      Left            =   1755
      TabIndex        =   3
      Text            =   "0000"
      Top             =   945
      Width           =   1230
   End
   Begin VB.TextBox txFile 
      Height          =   285
      Index           =   1
      Left            =   1755
      TabIndex        =   2
      Top             =   540
      Width           =   2580
   End
   Begin VB.TextBox txFile 
      Height          =   285
      Index           =   0
      Left            =   1755
      TabIndex        =   1
      Top             =   135
      Width           =   2580
   End
   Begin VB.Label Label3 
      Caption         =   "Second file starts at:"
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1395
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Split movie at frame:"
      Height          =   240
      Left            =   90
      TabIndex        =   8
      Top             =   990
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "2nd file name:"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   585
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "1st file name:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   1050
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------ SPLITS THE SUBTITLE FILE INTO 2 SEPARATE FILES. -----
'-               USEFUL FOR MOVIES ON 2 CDS

Option Explicit
Dim lSplitAt As Long
Dim lSplitFrame As Long

'--- Sets the split position (called by main window) ---
Sub SetPosition(ByVal SplitFrame, SplitPosition)
    Dim S As String
    frmSplit.txPos(0) = SplitFrame
    frmSplit.txPos(1) = Title(SplitPosition + 1).StartFrame - SplitFrame + 1
    S = BF.GetFileNameFromDir(sSubFileName)
    txFile(0) = Left(S, Len(S) - 4) + "-1" + Right(S, 4)
    txFile(1) = Left(S, Len(S) - 4) + "-2" + Right(S, 4)
    lSplitAt = SplitPosition
End Sub

Private Sub btCancel_Click()
    Me.Hide
End Sub

'--- Saves the 2 separate files to the disk ---
Private Sub btSave_Click()
    Dim sDir As String
    Dim tPart() As Subtitles, i As Long
    Dim lShift As Long
    
    'get the destination folder
    
    sDir = BF.BrowseFolder(Me.hwnd, "Select destination folder")
    If sDir = "" Then Exit Sub
    On Error GoTo 200
    
    lShift = Val(txPos(0)) - 1
    ReDim tPart(1 To lSplitAt)
    For i = 1 To lSplitAt
        tPart(i) = Title(i)
    Next i
    SaveTitles sDir + txFile(0), tPart()
    ReDim tPart(1 To LineCnt - lSplitAt)
    For i = lSplitAt + 1 To LineCnt
        tPart(i - lSplitAt).Text = Title(i).Text
        tPart(i - lSplitAt).StartFrame = Title(i).StartFrame - lShift
        tPart(i - lSplitAt).EndFrame = Title(i).EndFrame - lShift
    Next i
    SaveTitles sDir + txFile(1), tPart()
    Erase tPart
    Me.Hide
    MsgBox "Files succesfuly created into " + sDir, vbInformation
10
    Exit Sub
100
    Resume 10
200
    MsgBox Err.Description, vbCritical
    Resume 10
End Sub
