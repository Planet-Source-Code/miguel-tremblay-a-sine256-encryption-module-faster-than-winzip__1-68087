Attribute VB_Name = "StatsIt"
Option Explicit
Public Function GetTimeRemaining(FileL As Double, Chunk As Double, st As Long, i As Double) As Long
    GetTimeRemaining = Round(((Round(((FileL / (i * Chunk)) * (GetTickCount - st)))) - (GetTickCount - st)) / 1000) 'Return a value indicating seconds left
End Function
Public Function GetPercent(i As Double, segment As Double) As Integer
    GetPercent = Round((i * 100) / segment) 'Return a value between 0 and 100
End Function
Public Function GetSpeed(st As Long, Chunk As Double, i As Double) As String
    Dim xMb As Double
    xMb = GetTickCount - st
    xMb = xMb + 0.00001
    xMb = 1000 / xMb
    xMb = (i * Chunk) * xMb
    GetSpeed = GetFileSize(xMb) 'return string like "3.14Mb"
End Function
