VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SINE256cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const PI As Double = 3.14159265
Private y() As Double 'variable holding Y values in the graph : Y(x)  VERY IMPORTANT VARIABLE



Private Function Add256(current As Byte, add As Long) As Byte 'Add A number to a byte and return value between 0 and 255
    Add256 = ((current + add) Mod 256)
End Function
Private Function Min256(current As Byte, min As Long) As Byte 'Remove A number to a byte and return value between 0 and 255 (Inverse of add256)

    Min256 = ((current + 512) - (min Mod 256)) Mod 256 'Not that a big deal to reverse it
End Function
Public Sub SINE256Encrypt(input256() As Byte, output256() As Byte, password As String, PatternLen As Double)
    Dim xPass As Integer 'Variable holding the password value
    xPass = MakePasswordHash(password) ' Mess around with password string to make it integer and complicate to reproduce
    xPass = xPass * PI

    Dim x As Double 'the X of the GetYValue() function and the position in the text/bytearray
    ReDim output256(0 To UBound(input256)) 'Resize the output array to make it same size as input since my encryption give EXACT same size
    For x = 0 To UBound(input256) 'Let's start encrypting.. Loop from 0 to the length of the stuff to encrypt

        output256(x) = (input256(x) + y((x Mod (PatternLen - 1)) + 1) * xPass) Mod 256 'Heres the magic get Y value from X and add your xPass to it plus the current character being crypted then convert the number from 0 to 256
    Next
End Sub
Public Sub SINE256Decrypt(input256() As Byte, output256() As Byte, password As String, PatternLen As Double)
    Dim xPass As Integer 'Same thing again so we get same value to decrypt
    xPass = MakePasswordHash(password) 'Same thing again so we get same value to decrypt
    xPass = xPass * PI
    Dim MinNbr As Double
    Dim x As Double 'Same thing again so we get same value to decrypt
    ReDim output256(0 To UBound(input256)) 'Same thing again so we get same value to decrypt
    For x = 0 To UBound(input256) 'Same thing again so we get same value to decrypt


        
        output256(x) = ((input256(x) + 512) - ((xPass * y((x Mod (PatternLen - 1)) + 1)) Mod 256)) Mod 256 'Get Y from X and remove Xpass and the current encrypted chr. Convert the stuff in a 0-255 value (Inverse of Crypt)
    
    Next

End Sub
Public Function MakeGraph(decal As Double, start As Double, rangE As Double, Lenght As Double) As Double()
    ReDim y(1 To Lenght)
    Dim i As Double


    For i = 1 To Lenght
        
        
        y(i) = (Tan(i) * Exp(Log(i))) + Tan(Cos(i)) 'Get value to crypt or decrypt
        y(i) = y(i) * (Tan(i) * Log(i))
        y(i) = (y(i) + i)
        y(i) = Cos(y(i)) * Tan(i)
        y(i) = (Sin(y(i) / (PI * decal)) * rangE) + start

    If i Mod 16000 = 1 Then DoEvents
    Next
 
    MakeGraph = y()
End Function
