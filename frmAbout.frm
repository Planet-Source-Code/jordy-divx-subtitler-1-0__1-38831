VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Subtitler"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   1755
      TabIndex        =   3
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BugFull Software Inc., 2002"
      ForeColor       =   &H00000080&
      Height          =   330
      Index           =   2
      Left            =   45
      TabIndex        =   4
      Top             =   1035
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version "
      Height          =   330
      Index           =   1
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Coded by Jordy (chavdar_jordanov@yahoo.com)"
      ForeColor       =   &H00000080&
      Height          =   330
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   720
      Width           =   4605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "S U B T I T L E R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4515
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label2(1).Caption = "Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    
End Sub
