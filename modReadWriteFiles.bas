Attribute VB_Name = "modReadWriteFiles"
Option Explicit

Public Function ReadFile(ByVal sFilePath As String) As String
    Dim F As Long
    Dim S As String
    On Error Resume Next
    If FileLen(sFilePath) < 1 Then
        ReadFile = ""
        Exit Function
    Else
        F = FreeFile
        Open sFilePath For Binary Access Read As #F
        S = Space$(LOF(F))
        Get #F, , S
        Close #F
        ReadFile = S
        S = ""
    End If
End Function

Public Sub WriteFile(ByVal sFilePath As String, sString As String, Optional ByVal bAppend = False)
    Dim F As Long, L As Long
    On Error Resume Next
    F = FreeFile
    If Dir(sFilePath) <> "" Then Kill sFilePath
    Open sFilePath For Binary Access Write As #F
    If bAppend Then
        L = LOF(F)
        Seek #F, L + 1
    End If
    Put #F, , sString
    Close #F
End Sub

Public Function ReadFileInBytes(ByVal sFilePath As String)
    Dim F As Long
    Dim B()  As Byte
    
    On Error Resume Next
    If FileLen(sFilePath) < 1 Then
        Exit Function
    Else
        F = FreeFile
        Open sFilePath For Binary Access Read As #F
        ReDim B(LOF(F) - 1)
        Get #F, , B()
        Close #F
        ReadFileInBytes = B()
        Erase B
    End If
End Function

Public Sub WriteFileAsBytes(ByVal sFilePath As String, bBytes() As Byte)
    Dim F As Long
    On Error Resume Next
    F = FreeFile
    If Dir(sFilePath) <> "" Then Kill sFilePath
    Open sFilePath For Binary Access Write As #F
    Put #F, , bBytes()
    Close #F
End Sub
