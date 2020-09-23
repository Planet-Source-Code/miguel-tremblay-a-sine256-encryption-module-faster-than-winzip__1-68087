Attribute VB_Name = "PasswordHashing"
Option Explicit
Public Function MakePasswordHash(pass As String) As Double 'The function that processe a number from a string
    'The goal here is to mess around with the string and make it complicate
    'You can make 100 page of this if you wanted to
    'This value does not need to be reversed so you can make fun a long time
    If pass = "" Then pass = "None"
    Dim bPass() As Byte 'The password in byte (Byte array is very fast)
    Dim uHold() As Double
    Dim undValue As Double
    undValue = 3.14159265 'lets start with Pi value... Why not :/
    bPass = StrConv(pass, vbFromUnicode) 'convert the password string to byte array
    Dim i As Long
    ReDim uHold(0 To UBound(bPass)) 'Resize the other var
    For i = 0 To UBound(bPass) 'lets start messing around
        uHold(i) = (bPass(i) ^ 1.003) Mod 256687 'bleh
        undValue = (undValue + ((Len(pass) / 3.1416))) Mod 18168256
        undValue = (undValue + uHold(i)) Mod 16168256
        undValue = undValue / Oct(Len(pass))
        undValue = undValue * (Sin(i) + 1)
    Next
    'Well thats enuff for now
    'Lets test this out ;)
    MakePasswordHash = Round(undValue * (1048 / 3)) Mod 1048 'Make it a value between 0 and 255 .. i dont know why, just for fun ;)
End Function
