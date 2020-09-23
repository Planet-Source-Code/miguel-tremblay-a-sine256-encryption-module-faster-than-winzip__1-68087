Attribute VB_Name = "Vars"
Option Explicit
Public password As String 'Var holding the password string
Public decal As Double 'Var holding the decal double
Public start As Double 'Var holding the start double
Public rangE As Double 'Var holding the range double
Public PatternLen As Double 'Var holding until when X can go until it repeat the algorithm
Public Chunk As Double 'Buffer size per segment
Public cl As New SINE256cls
