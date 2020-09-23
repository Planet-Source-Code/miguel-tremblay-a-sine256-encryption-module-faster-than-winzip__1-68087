Attribute VB_Name = "GaphIt"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long 'Just for stats
Private Const PI As Double = 3.14159265
Public Const BIGNUMBER_32 As Double = (2 ^ 32)

'THERES NOTHING REALLY IMPORTANT HERE
'JUST TO VISUALISE THE ALGORITHM
Public Function ChartIt(Chart As MSChart, ArrayX() As Byte)
    Dim i As Long
    Dim a As Long
    a = UBound(ArrayX)
    Dim k() As Long
    ReDim k(0 To a)
    For i = 0 To UBound(ArrayX)
        k(i) = ArrayX(i) 'Convert bytes to long
    Next
    Chart.ChartData = k 'Draw datas
    For i = 0 To UBound(ArrayX) 'Change graph color
        Chart.Plot.SeriesCollection(i + 1).DataPoints(-1).Brush.FillColor.Red = 0
        Chart.Plot.SeriesCollection(i + 1).DataPoints(-1).Brush.FillColor.Blue = 150
        Chart.Plot.SeriesCollection(i + 1).DataPoints(-1).Brush.FillColor.Green = 70
    Next
End Function

Public Function ChartIt2(Chart As MSChart, y() As Double)
    Dim k As Long
    Dim R(1 To 5000) As Double
    Dim Buff As Double
    For k = 1 To 5000
        Buff = Round(((UBound(y)) / 5000) * k)
        If Buff = 0 Then Buff = 1
        R(k) = y(Buff)
    Next
    Chart.ChartData = R()
    For k = 1 To 5000
        Chart.Plot.SeriesCollection(k).DataPoints(-1).Brush.FillColor.Red = 0
        Chart.Plot.SeriesCollection(k).DataPoints(-1).Brush.FillColor.Blue = 150
        Chart.Plot.SeriesCollection(k).DataPoints(-1).Brush.FillColor.Green = 70
    Next
End Function

