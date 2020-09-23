Attribute VB_Name = "modColor"
'*****************************************************************
'*              Color processing subroutines                     *
'*              written by Chavdar Yordanov, 04.2001             *
'*              Email: chavdar_jordanov@yahoo.com                *
'*              Please, don't remove this title!                 *
'*****************************************************************

Option Explicit

Public RGBs() As String
Public SafeCol(224) As Long
Public iColorDepth As Integer

'dividers for GetColorByte function
Public Const clr24Bit = 1
Public Const clr16Bit = 8
Public Const clrWebSafe = 51

Public Function HexToLong(sHexColor As String) As Long
    Dim lCol As Long, i, N
    If Left(sHexColor, 1) = Chr(34) Then sHexColor = Mid(sHexColor, 2, Len(sHexColor) - 2)
    If Left(sHexColor, 1) = "#" Then sHexColor = Mid(sHexColor, 2)
    sHexColor = UCase(sHexColor)
    lCol = 0: N = 0
    For i = Len(sHexColor) - 1 To 1 Step -2
        lCol = lCol + Dec(Mid(sHexColor, i, 2)) * 256 ^ N
        N = N + 1
    Next i
    HexToLong = lCol
End Function

Public Function RgbToLong(sRgbColor As String) As Long
    Dim vCol, i, lCol As Long, N
    vCol = Split(sRgbColor, ",")
    For i = LBound(vCol) To UBound(vCol)
        lCol = lCol + Val(vCol(i)) * 256 ^ N
        N = N + 1
    Next i
    RgbToLong = lCol
End Function

Public Function Dec(ByVal sHex As String) As Long 'Converts Hex to Decimal
    Const HVal = "0123456789ABCDEF"
    Dim iPos As Byte, i As Integer, lDec As Long
    Dim L As Integer, X As Byte
    L = Len(sHex)
    If L > 255 Then Exit Function
    lDec = 0
    For i = L To 1 Step -1
        X = InStr(1, HVal, Mid(sHex, i, 1), vbTextCompare)
        If X = 0 Then Exit Function Else X = X - 1
        lDec = lDec + X * 16 ^ (L - i)
    Next i
    Dec = lDec
End Function

Public Function Invert(ByVal iCol As Long) As Long
    Dim bCol() As Byte    'Byte values
    SplitIntoBytes iCol, 3, bCol()
    Invert = RGB(255 - bCol(1), 255 - bCol(2), 255 - bCol(3))
End Function

Public Sub SplitIntoBytes(ByVal lNumber As Long, bSize As Byte, ByRef bBytes() As Byte, Optional bRedim = True)
    Dim i As Long
    Dim KF As Long
    If bRedim Then ReDim bBytes(1 To bSize)
    For i = bSize To 1 Step -1
        KF = 256 ^ (i - 1)
        bBytes(i) = lNumber \ KF
        lNumber = lNumber - bBytes(i) * KF
    Next i
End Sub


Public Function CalcColorDepth(lColor As Long) As Long
    CalcColorDepth = GetColorByte(lColor And &HFF&)
    CalcColorDepth = CalcColorDepth + CLng(GetColorByte((lColor And &HFF00&) \ &H100&)) * 256
    CalcColorDepth = CalcColorDepth + CLng(GetColorByte((lColor And &HFF0000) \ &H10000)) * 65536
End Function

Public Function GetColorByte(bCol As Long) As Byte
    Dim z As Long
    z = iColorDepth * Int((bCol - (bCol > (255 - iColorDepth))) / iColorDepth)
    GetColorByte = CByte(z + (z > 255))
End Function

Public Function LongToHex(lCol As Long) As String
    Dim B() As Byte
    Dim i
    Dim sRgb As String
    SplitIntoBytes lCol, 3, B()
    For i = 3 To 1 Step -1
        sRgb = sRgb + Format(Hex(B(i)), "00")
    Next i
    LongToHex = sRgb
End Function

Function Convert2Long(ByVal sColor As String) As Long
    sColor = Replace(sColor, "$", "")
    sColor = Right(sColor, 6)
    Convert2Long = HexToLong(sColor)
End Function

Function Convert2MDVD(ByVal lColor As Long) As String
    Dim sHex As String
    Convert2MDVD = "$" + LongToHex(lColor)
End Function
