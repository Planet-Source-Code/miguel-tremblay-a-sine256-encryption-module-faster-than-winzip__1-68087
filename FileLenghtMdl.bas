Attribute VB_Name = "FileLenghtMdl"
Option Explicit
Public Function LengthOfFile(ByVal Filename As String) As Double
    LengthOfFile = dblUnsigned(FileLen(Filename)) 'Call the FileLen Function and convert it To an unsigned value
End Function
Function lngSigned(ByVal dblUnsigned As Double) As Long
    If dblUnsigned <= &H7FFFFFFF Then
        lngSigned = dblUnsigned
    Else
        lngSigned = CLng(dblUnsigned - BIGNUMBER_32)
    End If
End Function

Public Function GetFileSize(FileLength As Double) As String
    Dim iSizeMB As Double
    On Error GoTo filelenerr
    iSizeMB = Round((FileLength / 1024) / 1024, 2)
    GetFileSize = iSizeMB & "Mb"
    Exit Function
filelenerr:
    GetFileSize = -1
End Function
Function dblUnsigned(ByVal lngSigned As Long) As Double
    If lngSigned >= 0 Then
        dblUnsigned = lngSigned
    Else
        dblUnsigned = BIGNUMBER_32 + lngSigned
    End If
End Function
