VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5130
      TabIndex        =   5
      Top             =   585
      Width           =   1275
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5130
      TabIndex        =   4
      Top             =   135
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5055
      Begin VB.CheckBox chMeasure 
         Caption         =   "Measure frame rate on movie open"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   990
         Width           =   3210
      End
      Begin VB.TextBox txSet 
         Height          =   285
         Index           =   1
         Left            =   1980
         TabIndex        =   2
         Text            =   "20"
         Top             =   585
         Width           =   690
      End
      Begin VB.TextBox txSet 
         Height          =   285
         Index           =   0
         Left            =   1980
         TabIndex        =   1
         Text            =   "500"
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "seconds"
         Height          =   195
         Index           =   1
         Left            =   2745
         TabIndex        =   9
         Top             =   630
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Measure frame rate for:"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   630
         Width           =   1770
      End
      Begin VB.Label Label2 
         Caption         =   "milliseconds"
         Height          =   195
         Index           =   0
         Left            =   2745
         TabIndex        =   7
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Screen update interval:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--- SETTINGS DIALOG ----
Option Explicit

Private Sub btCancel_Click()
    Me.Hide
End Sub
'--- saves and applies settings ---
Private Sub btOK_Click()
    lUpdateTime = Val(txSet(0))
    lMeasureTime = Val(txSet(1)) * 1000
    bAutoMeasure = chMeasure.Value = 1
    SaveSettings
    frmSub.InitSettings
    Me.Hide
End Sub

Private Sub Form_Load()
    txSet(0) = CStr(lUpdateTime)
    txSet(1) = CStr(lMeasureTime / 1000)
    chMeasure.Value = Abs(bAutoMeasure)
End Sub
