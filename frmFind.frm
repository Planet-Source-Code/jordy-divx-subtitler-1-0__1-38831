VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Find text"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   3285
      TabIndex        =   10
      Top             =   1575
      Width           =   1185
   End
   Begin VB.CheckBox chReplace 
      Caption         =   "Preserve original case when replacing"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   1620
      Width           =   3030
   End
   Begin VB.CheckBox chReplace 
      Caption         =   "Whole words"
      Height          =   195
      Index           =   1
      Left            =   1350
      TabIndex        =   4
      Top             =   1350
      Width           =   1275
   End
   Begin VB.CheckBox chReplace 
      Caption         =   "Match case"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   1350
      Width           =   1185
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   1905
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btFind 
      Caption         =   "Replace &All"
      Height          =   330
      Index           =   3
      Left            =   3285
      TabIndex        =   9
      Top             =   1125
      Width           =   1185
   End
   Begin VB.CommandButton btFind 
      Caption         =   "&Replace"
      Height          =   330
      Index           =   2
      Left            =   3285
      TabIndex        =   8
      Top             =   765
      Width           =   1185
   End
   Begin VB.CommandButton btFind 
      Caption         =   "Find &next"
      Height          =   330
      Index           =   1
      Left            =   3285
      TabIndex        =   7
      Top             =   405
      Width           =   1185
   End
   Begin VB.CommandButton btFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   3285
      TabIndex        =   6
      Top             =   45
      Width           =   1185
   End
   Begin VB.TextBox txFind 
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   2850
   End
   Begin VB.TextBox txFind 
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2850
   End
   Begin VB.Label Label1 
      Caption         =   "Replace with:"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Find text:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   825
   End
   Begin VB.Menu mnPop 
      Caption         =   "mnPop"
      Visible         =   0   'False
      Begin VB.Menu mnFrom 
         Caption         =   "From this point on"
         Index           =   0
      End
      Begin VB.Menu mnFrom 
         Caption         =   "Until this point"
         Index           =   1
      End
      Begin VB.Menu mnFrom 
         Caption         =   "Entire movie"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------- SEARCH AND REPLACE DIALOG -------------

Option Explicit
Dim lCnt As Long

Private Sub Form_Load()
    lRowFound = 0
End Sub

Private Sub btClose_Click()
    Me.Hide
End Sub

Private Sub btFind_Click(Index As Integer)
    Dim iCompareMode As Integer
    If chReplace(0).Value Then iCompareMode = vbBinaryCompare Else iCompareMode = vbTextCompare
    Select Case Index
    Case 0 'find
        FindTitle txFind(0), True, iCompareMode
    Case 1 'find next
        FindTitle txFind(0), False, iCompareMode
    Case 2 'replace
        ReplaceTitle txFind(0), txFind(1), False, iCompareMode, chReplace(1).Value = 1, chReplace(2).Value = 1
    Case 3 'replace all
        ReplaceTitle txFind(0), txFind(1), True, iCompareMode, chReplace(1).Value = 1, chReplace(2).Value = 1
    End Select
End Sub

'--- Searches the subtitles and replaces strings ---
Sub ReplaceTitle(sFind As String, sReplace As String, bAll As Boolean, iCompareMode As Integer, bWhole As Boolean, bPreserve As Boolean)
    Dim lFound As Long
    If sFind = "" Then Exit Sub
    
    If bAll Then
        If MsgBox("Replace ALL occurrences of '" + sFind + "' with '" + sReplace + "'?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
        lCnt = 0
        Do
            lFound = FindTitle(sFind, False, iCompareMode)
            If lFound > 0 Then
                lCnt = lCnt + ChangeTitle(lFound, sFind, sReplace, iCompareMode, bWhole, bPreserve)
                Title(lFound).Text = Replace(Title(lFound).Text, sFind, sReplace, 1, -1, vbTextCompare)
                frmSub.GRID.TextMatrix(lFound, 3) = Title(lFound).Text
            End If
        Loop Until lFound = 0
        Stat.SimpleText = CStr(lCnt) + " replacements made. "
    Else
        If lRowFound = 0 Then
            lFound = FindTitle(sFind, False, iCompareMode)
        Else
            lFound = lRowFound
        End If
        If lFound > 0 Then
            lCnt = lCnt + ChangeTitle(lFound, sFind, sReplace, iCompareMode, bWhole, bPreserve)
        End If
    End If
End Sub

'--- Searches the subtitles for a string and returns its index ---
Function FindTitle(sFind As String, bFindFirst As Boolean, iCompareMode As Integer) As Long
    Dim i As Long
    If bFindFirst Then
        lCnt = 0
        lRowFound = 0
    End If
    For i = lRowFound + 1 To LineCnt
        If InStr(1, Title(i).Text, sFind, iCompareMode) > 0 Then
            Stat.SimpleText = "Found at " + CStr(i)
            frmSub.GRID.TopRow = i
            lRowFound = i
            FindTitle = i
            Exit Function
        End If
    Next i
    Stat.SimpleText = "Not found!"
    FindTitle = 0
End Function

'--- Replaces occurences of a string with a new string, preserves case ---
Function ChangeTitle(lFound As Long, sFind As String, sReplace As String, iCompareMode As Integer, bWhole As Boolean, bPreserve As Boolean) As Long
    Dim X As Integer, S As String
    Dim CheckLine As String, C As String
    Dim sOldText As String, sNewText As String
    Dim Counter As Long
    
    CheckLine = " ,./-+!@?;:$#%^&*'" + Chr(34)
    S = Title(lFound).Text
    Do
NextTry:
        X = InStr(X + 1, S, sFind, iCompareMode)
        If X = 0 Then Exit Do
        If bWhole Then
            C = Mid(S, X + Len(sFind) + 1)
            If InStr(CheckLine, C) Then GoTo NextTry
        End If
        sNewText = sReplace
        If bPreserve Then
            sOldText = Mid(S, X, Len(sFind))
            If UCase(sOldText) = sOldText Then
                sNewText = UCase(sReplace)
            ElseIf LCase(sOldText) = sOldText Then
                sNewText = LCase(sReplace)
            ElseIf UCase(Left(sOldText, 1)) = Left(sOldText, 1) Then
                sNewText = BF.CapitalizeString(sReplace)
            End If
        End If
        S = Left(S, X - 1) + sNewText + Mid(S, X + Len(sFind))
        X = X + Len(sReplace)
        Counter = Counter + 1
    Loop
    Title(lFound).Text = S
    frmSub.GRID.TextMatrix(lFound, 3) = S
    ChangeTitle = Counter
End Function
