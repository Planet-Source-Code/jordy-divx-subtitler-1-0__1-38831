VERSION 5.00
Begin VB.Form frmAction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Re-calculate frames"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmAction.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2025
      Width           =   1275
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   45
      TabIndex        =   10
      Top             =   2025
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.CheckBox chRecalc 
         Caption         =   "Recalculate frames for the rest of the movie"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1575
         Value           =   1  'Checked
         Width           =   3795
      End
      Begin VB.CommandButton btGetFrame 
         Caption         =   "Get from movie"
         Height          =   285
         Left            =   2835
         TabIndex        =   3
         Top             =   225
         Width           =   1230
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Add"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   705
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change value to:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1620
      End
      Begin VB.ComboBox cbFPS 
         Height          =   315
         Index           =   0
         ItemData        =   "frmAction.frx":014A
         Left            =   2565
         List            =   "frmAction.frx":0163
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   675
         Width           =   825
      End
      Begin VB.TextBox txSec 
         Height          =   285
         Left            =   945
         TabIndex        =   5
         Text            =   "00"
         Top             =   690
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "0000"
         Top             =   225
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Change entire movie FPS to:"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1125
         Width           =   2355
      End
      Begin VB.ComboBox cbFPS 
         Height          =   315
         Index           =   1
         ItemData        =   "frmAction.frx":0185
         Left            =   2565
         List            =   "frmAction.frx":019E
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1035
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FPS"
         Height          =   195
         Index           =   0
         Left            =   3480
         TabIndex        =   14
         Top             =   705
         Width           =   300
      End
      Begin VB.Label Label4 
         Caption         =   "seconds at "
         Height          =   225
         Left            =   1635
         TabIndex        =   13
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "FPS"
         Height          =   195
         Index           =   1
         Left            =   3465
         TabIndex        =   12
         Top             =   1080
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btCancel_Click()
    Me.Hide
End Sub

Private Sub btGetFrame_Click()
    frmSub.MMPlayer.DisplayMode = mpFrames
    Text1.Text = CStr(frmSub.MMPlayer.CurrentPosition)
End Sub

Private Sub btOK_Click()
    Dim iValue, Change, FirstIndex, iFPS
    If Option1(0).Value Then
        If bTitleStart Then
            Change = Val(Text1.Text) - Title(lSelStart).StartFrame
        Else
            Change = Val(Text1.Text) - Title(lSelStart).EndFrame
        End If
    ElseIf Option1(1).Value Then
        iFPS = Val(cbFPS(0).Text)
        iValue = Val(txSec.Text)
        Change = Int(iValue * iFPS + 0.5)
    Else
        ChangeFPS gFPS, Val(cbFPS(1).Text)
        frmSub.cbFPS.Text = cbFPS(1).Text
    End If
    If chRecalc.Value = 1 Then
        ChangeIndex lSelStart, bTitleStart, Change
    Else
        ChangeTitle lSelStart, bTitleStart, Change
    End If
    frmSub.RefreshTitles
    Me.Hide
End Sub

Private Sub Form_Load()
    cbFPS(0).Text = CStr(gFPS)
    cbFPS(1).Text = "25"
End Sub

Private Sub Option1_Click(Index As Integer)
    chRecalc.Visible = Index < 2
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Not IsNumeric(Text1.Text) Then
        MsgBox "Please, enter a valid number into this field!", vbExclamation
        Cancel = True
    End If
End Sub
