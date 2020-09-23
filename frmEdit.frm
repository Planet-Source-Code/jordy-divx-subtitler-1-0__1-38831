VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit subtitle"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CCol 
      Left            =   3105
      Top             =   3465
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   16777215
   End
   Begin VB.TextBox txTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   4
      Left            =   180
      TabIndex        =   4
      Top             =   1485
      Width           =   5595
   End
   Begin VB.TextBox txTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   1125
      Width           =   5595
   End
   Begin VB.TextBox txTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   765
      Width           =   5595
   End
   Begin VB.TextBox txTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   405
      Width           =   5595
   End
   Begin VB.TextBox txTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   45
      Width           =   5595
   End
   Begin VB.Frame Frame1 
      Caption         =   "Format"
      Height          =   1005
      Left            =   90
      TabIndex        =   22
      Top             =   2340
      Width           =   5730
      Begin VB.CheckBox chALL 
         Caption         =   "Apply to all"
         Height          =   195
         Left            =   135
         TabIndex        =   14
         Top             =   675
         Value           =   1  'Checked
         Width           =   2085
      End
      Begin VB.TextBox txSize 
         Height          =   285
         Left            =   4905
         TabIndex        =   13
         Text            =   " "
         Top             =   315
         Width           =   555
      End
      Begin VB.CommandButton btCol 
         Caption         =   "Clear"
         Height          =   285
         Index           =   1
         Left            =   3735
         TabIndex        =   12
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton btCol 
         Caption         =   "Set"
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   11
         Top             =   315
         Width           =   690
      End
      Begin VB.CheckBox chFmt 
         Height          =   285
         Index           =   2
         Left            =   765
         Picture         =   "frmEdit.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chFmt 
         Height          =   285
         Index           =   1
         Left            =   450
         Picture         =   "frmEdit.frx":09CC
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chFmt 
         Height          =   285
         Index           =   0
         Left            =   135
         Picture         =   "frmEdit.frx":0ACE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label Label6 
         Caption         =   "Size:"
         Height          =   240
         Left            =   4455
         TabIndex        =   24
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lbCol 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<default>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2025
         TabIndex        =   18
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "Font color:"
         Height          =   240
         Left            =   1260
         TabIndex        =   23
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.ComboBox cbTime 
      Height          =   315
      ItemData        =   "frmEdit.frx":0BD0
      Left            =   2610
      List            =   "frmEdit.frx":0BF2
      TabIndex        =   6
      Text            =   "3.0"
      Top             =   1980
      Width           =   735
   End
   Begin VB.TextBox txEnd 
      Height          =   285
      Left            =   4635
      TabIndex        =   7
      Top             =   1980
      Width           =   1185
   End
   Begin VB.TextBox txStart 
      Height          =   285
      Left            =   630
      TabIndex        =   5
      Top             =   1980
      Width           =   1005
   End
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4500
      TabIndex        =   16
      Top             =   3375
      Width           =   1365
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   90
      TabIndex        =   15
      Top             =   3375
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "sec."
      Height          =   195
      Left            =   3375
      TabIndex        =   21
      Top             =   2025
      Width           =   330
   End
   Begin VB.Label Label3 
      Caption         =   "Show for"
      Height          =   195
      Left            =   1845
      TabIndex        =   20
      Top             =   2025
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Ends:"
      Height          =   240
      Left            =   4140
      TabIndex        =   19
      Top             =   2025
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Starts:"
      Height          =   240
      Left            =   90
      TabIndex        =   17
      Top             =   2025
      Width           =   555
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bInsertMode As Boolean
Dim sFormat(5) As String
Dim CurrRow As Integer

Private Sub btCancel_Click()
    Me.Hide
End Sub

Private Sub btCol_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case 0 'set color
        CCol.Flags = 3
        CCol.Color = lbCol.BackColor
        CCol.ShowColor
        If chALL.Value = 1 Then
            For i = 0 To txTitle.Count - 1
                txTitle(i).ForeColor = CCol.Color
            Next i
        Else
            txTitle(CurrRow).ForeColor = CCol.Color
        End If
    Case 1 'clear color
        If chALL.Value = 1 Then
            For i = 0 To txTitle.Count - 1
                txTitle(i).ForeColor = vbWhite
            Next i
        Else
            txTitle(CurrRow).ForeColor = vbWhite
        End If
    End Select
    txTitle(CurrRow).SetFocus
End Sub

Private Sub btOK_Click()
    If bInsertMode Then InsertTitle lSelStart
    Title(lSelStart).Text = CompileTitle
    Title(lSelStart).StartFrame = Val(Trim(txStart.Text))
    Title(lSelStart).EndFrame = Val(Trim(txEnd.Text))
    frmSub.GoFrame lSelStart
    frmSub.RefreshTitles
    bChanged = True
    Me.Hide
End Sub

Sub SetTitle(ByVal lIndex As Long)
    Dim H, i As Byte
    'On Error Resume Next
    bInsertMode = False
    ClearTitles
    FormatBoxes Title(lIndex).Text

    txStart = CStr(Title(lIndex).StartFrame)
    txEnd = CStr(Title(lIndex).EndFrame)
    cbTime.Text = Format(((Title(lIndex).EndFrame) - Title(lIndex).StartFrame) / gFPS, "0.0")
End Sub

Public Sub InsertAt(ByVal lIndex As Long, Optional StartFrame = -1, Optional EndFrame = -1)
    On Error Resume Next
    ClearTitles
    If StartFrame < 0 Then
        txStart = CStr(Title(lIndex - 1).EndFrame + 2)
    Else
        txStart = CStr(StartFrame)
    End If
    If EndFrame < 0 Then
        txEnd = Format(Title(lIndex - 1).EndFrame + 2 + Val(cbTime.Text) * gFPS, "0")
    Else
        txEnd = CStr(EndFrame)
    End If
    bInsertMode = True
    Me.Show 1
End Sub
Private Sub cbTime_Click()
    Recalc
End Sub

Sub Recalc()
    txEnd = Format(Val(txStart) + 2 + Val(cbTime.Text) * gFPS, "0")
End Sub

Private Sub cbTime_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) Then
        cbTime.Text = cbTime.Text + Chr(KeyAscii)
        KeyAscii = 0
        Recalc
    End If
End Sub


Private Sub chFmt_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case 0
        If chALL.Value = 1 Then
            For i = 0 To txTitle.Count - 1
                txTitle(i).FontBold = -chFmt(0).Value
            Next i
        Else
            txTitle(CurrRow).FontBold = -chFmt(0).Value
        End If
    Case 1
        If chALL.Value = 1 Then
            For i = 0 To txTitle.Count - 1
                txTitle(i).FontItalic = -chFmt(1).Value
            Next i
        Else
            txTitle(CurrRow).FontItalic = -chFmt(1).Value
        End If
    Case 2
        If chALL.Value = 1 Then
            For i = 0 To txTitle.Count - 1
                txTitle(i).FontUnderline = -chFmt(2).Value
            Next i
        Else
            txTitle(CurrRow).FontUnderline = -chFmt(2).Value
        End If
    End Select
End Sub

Private Sub Form_Activate()
    txTitle(0).SetFocus
End Sub

Private Sub txEnd_Change()
    cbTime.Text = Format((Val(txEnd.Text) - Val(txStart.Text)) / gFPS, "0.0")
End Sub

Private Sub txSize_Change()
    txTitle(CurrRow).Tag = txSize.Text
End Sub

Private Sub txSize_GotFocus()
    With txSize
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txStart_Change()
    Recalc
End Sub

Sub FormatBoxes(sTitleText As String)
    Dim i As Integer, S As String, sFmt As String, j As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim H, G
    G = Split(sTitleText, "|")
    ClearTitles
    For i = 0 To txTitle.Count - 1
        If i <= UBound(G) Then
            H = Split(G(i), "}")
            If UBound(H) >= 0 Then
                For j = 0 To UBound(H) - 1
                    SetFormat i, H(j)
                Next j
                txTitle(i).Text = H(j)
            End If
        End If
    Next i
End Sub

Function CompileTitle() As String
    Dim i As Integer, j As Integer
    Dim S As String
    Dim sFmt() As String
    Dim bFmt() As Boolean
    Dim iRowCnt As Integer
    Dim Ctrl, sAll As String
    Dim sRow() As String
    Ctrl = Array("y", "c", "s")
    ReDim sFmt(2, 0 To txTitle.Count - 1), bFmt(0 To txTitle.Count - 1)
    For i = 0 To txTitle.Count - 1
        If Trim(txTitle(i).Text) <> "" Then
            If txTitle(i).FontBold Then sFmt(0, i) = sFmt(0, i) + "b"
            If txTitle(i).FontItalic Then sFmt(0, i) = sFmt(0, i) + "i"
            If txTitle(i).FontUnderline Then sFmt(0, i) = sFmt(0, i) + "u"
            If txTitle(i).ForeColor <> vbWhite Then sFmt(1, i) = CStr(Convert2MDVD(txTitle(i).ForeColor))
            sFmt(2, i) = txTitle(i).Tag
            iRowCnt = iRowCnt + 1
        End If
    Next i
    If iRowCnt < 1 Then Exit Function
    sAll = ""
    ReDim sRow(0 To iRowCnt - 1)
    For i = 1 To iRowCnt - 1
        For j = 0 To 2
            If sFmt(j, i) <> sFmt(j, 0) Then bFmt(j) = True
        Next j
    Next i
    For j = 0 To 2
        If Not bFmt(j) And sFmt(j, 0) <> "" Then sAll = sAll + "{" + UCase(Ctrl(j)) + ":" + sFmt(j, 0) + "}"
    Next j
    
    For i = 0 To iRowCnt - 1
        For j = 0 To 2
            If bFmt(j) And sFmt(j, i) <> "" Then
                sRow(i) = sRow(i) + "{" + LCase(Ctrl(j)) + ":" + sFmt(j, i) + "}"
            End If
        Next j
        S = S + sRow(i) + txTitle(i).Text + "|"
    Next i
    CompileTitle = sAll + Left(S, Len(S) - 1)
End Function
Sub SetFormat(ByVal Index As Integer, ByVal sFmt As String)
    Dim sCtrl As String, i As Integer, sParam As String
    sCtrl = Mid(sFmt, 2, 2)
    sParam = Mid(sFmt, 4, Len(sFmt) - 1)
    Select Case sCtrl
    Case "y:"
        txTitle(Index).FontItalic = InStr(4, sFmt, "i", vbTextCompare) <> 0
        txTitle(Index).FontBold = InStr(4, sFmt, "b", vbTextCompare) <> 0
        txTitle(Index).FontUnderline = InStr(4, sFmt, "u", vbTextCompare) <> 0
    Case "Y:"
        For i = 0 To txTitle.Count - 1
            txTitle(i).FontItalic = InStr(4, sFmt, "i", vbTextCompare) <> 0
            txTitle(i).FontBold = InStr(4, sFmt, "b", vbTextCompare) <> 0
            txTitle(i).FontUnderline = InStr(4, sFmt, "u", vbTextCompare) <> 0
        Next i
    Case "s:"
        txTitle(Index).Tag = sParam
    Case "S:"
        For i = 0 To txTitle.Count - 1
            txTitle(i).Tag = sParam
        Next i
'    Case "f:"
'        txTitle(Index).FontName = sParam
'    Case "F:"
'        For i = 0 To txTitle.Count - 1
'            txTitle(i).FontName = sParam
'        Next i
    Case "c:"
        txTitle(Index).ForeColor = Convert2Long(sParam)
    Case "C:"
        For i = 0 To txTitle.Count - 1
            txTitle(i).ForeColor = Convert2Long(sParam)
        Next i
    End Select
End Sub

Sub ClearTitles()
    Dim i As Integer
    For i = 0 To txTitle.Count - 1
        txTitle(i).Text = ""
        txTitle(i).ForeColor = vbWhite
        txTitle(i).FontItalic = False
        txTitle(i).FontBold = False
        txTitle(i).FontName = "Arial Cyr"
        txTitle(i).Tag = ""
    Next i
    lbCol.BackColor = vbWhite
    lbCol.Caption = "<default>"
End Sub

Private Sub txTitle_GotFocus(Index As Integer)
    CurrRow = Index
    chFmt(0).Value = Abs(txTitle(Index).FontBold)
    chFmt(1).Value = Abs(txTitle(Index).FontItalic)
    chFmt(2).Value = Abs(txTitle(Index).FontUnderline)
    txSize.Text = txTitle(Index).Tag
    lbCol.BackColor = txTitle(Index).ForeColor
    If lbCol.BackColor = vbWhite Then lbCol.Caption = "<default>" Else lbCol.Caption = Convert2MDVD(lbCol.BackColor)
End Sub
