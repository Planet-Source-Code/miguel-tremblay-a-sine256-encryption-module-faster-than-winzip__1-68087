VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SINE256 Visualisation And Sample Console"
   ClientHeight    =   7770
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   13470
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart CryptChart 
      Height          =   3015
      Left            =   6240
      OleObjectBlob   =   "MainFrm.frx":030A
      TabIndex        =   6
      Top             =   5160
      Width           =   7335
   End
   Begin MSChart20Lib.MSChart TextChart 
      Height          =   3135
      Left            =   6240
      OleObjectBlob   =   "MainFrm.frx":2C65
      TabIndex        =   5
      Top             =   2280
      Width           =   7335
   End
   Begin MSChart20Lib.MSChart AlgoChart 
      Height          =   3135
      Left            =   6240
      OleObjectBlob   =   "MainFrm.frx":55C1
      TabIndex        =   7
      Top             =   -240
      Width           =   7335
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   6240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Statistic (Time in millisecond for each action)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   6015
      Begin VB.Label Label7 
         Caption         =   "String to byte convertion :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Encrypt/Decrypt time :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Byte to String convertion :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label10 
         Caption         =   "Display Data :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   17
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label StatsLbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label StatsLbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label StatsLbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   14
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label StatsLbl 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   13
         Top             =   1800
         Width           =   3495
      End
   End
   Begin VB.CommandButton DecryptFile 
      Caption         =   "Decrypt File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton EncryptFile 
      Caption         =   "Encrypt File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton SwapCmd 
      Caption         =   "Send in Encrypt TextBox"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   2295
   End
   Begin VB.CommandButton Decrypt 
      Caption         =   "Decrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox t2 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MainFrm.frx":7EFB
   End
   Begin RichTextLib.RichTextBox t1 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"MainFrm.frx":7F7D
   End
   Begin VB.CommandButton Encrypt 
      Caption         =   "Encrypt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Character Lenght :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   24
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label ResLenLbl 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Top             =   7080
      Width           =   4095
   End
   Begin VB.Label ChrLenLbl 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label11 
      Caption         =   "Character Lenght :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "*there is some Chr you cant copy paste"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Result (Will save to a file if Encrypt File is used) :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Width           =   6015
   End
   Begin VB.Menu SettingCmd 
      Caption         =   "Setting"
   End
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Let me explain what the graph does!
'This is the base algorithm where X value is the position in the string to encrypt
'This 100% came out of my mind and i personnaly never saw any encryption using this method
'If you know whats a Sin() and how it look in a graph you can read the rest
'decal is the X interval between each time Y hit 0 or whatever is your range and start value
'*Note Higher decal = more chance of character repetition
'Start is the Y starting position
'Range is the Minimum and maximum Y can reach ex: Start - Range is minimum and Start + Range is maximum
'Dont use higher range than start position and dont make it get past 256
'USING A 0 RANGE DOES NOT MEAN THE ENCRYPTED VERSION WILL BE 1 KIND OF CHR
'IN THE ENCRYPTION THE Y VALUE IS MODIFIED



Private Sub DecryptFile_Click()
Encrypt.Enabled = False 'Make button impossible to use
Decrypt.Enabled = False 'Make button impossible to use
EncryptFile.Enabled = False 'Make button impossible to use
DecryptFile.Enabled = False 'Make button impossible to use
DoEvents
    'Master variables initializing
    Dim Path As String 'Var holding path of file to encrypt or decrypt
    Dim SavePath As String 'Var holding the save path
    Dim st As Long 'Var holding value of GetTickCount when action start (Statistics)
    Dim Crypt() As Byte 'Variable getting file data (This var get resized and get different size of the file for optimization)
    Dim Res() As Byte 'This is same as Crypt() but its the datas once encrypted (Array is gonna be EXACT SAME SIZE)
    'Show Unused in stats frame
    StatsLbl(0).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(1).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(2).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(3).Caption = "Unused" 'Show that those kind of stats arent used with files
    '###Get file to encrypt###
    Dlg.DialogTitle = "Select File to Encrypt/Decrypt"
    Dlg.ShowOpen
    Path = Dlg.Filename
    '#########################
    '####Get path to save#####
    Dlg.DialogTitle = "Enter path to save file"
    Dlg.ShowOpen
    SavePath = Dlg.Filename
    '#########################
    If Not SavePath = "" And Not Path = "" Then
    
        Dim xVal() As Byte 'Buffer variable
        Dim segment As Double 'Number of segment Fix(Size / buffer) + 1
        Dim FileL As Double 'Length of the file
        FileL = LengthOfFile(Path) 'Set length of file
        segment = Fix(FileL / Chunk) + 1 'Set Number of segment
        Dim i As Double 'Variable used to loop (also the position in the segment)
        st = GetTickCount ' sT is the starting gettickcount value . Used for statistics
        Dim Freenum As Integer 'File number of the file to encrypt/decrypt
        Freenum = FreeFile 'Initialize file number
        Dim Freenum2 As Integer 'File number of the path to save
        Dim Str As String
        Open Path For Binary Access Read As #Freenum 'Open the first path before giving file number to save path (otherwise you will get same value for both)
        Freenum2 = FreeFile 'Now you can give him a number
        Open SavePath For Binary Access Write As #Freenum2 'Initialize the save path
        For i = 1 To segment 'Loop in the file buffers (Segment)
            Dim PercentX As Double 'Variable holding percent value (Statistics)
            Dim TimeRemain As Double 'Variable holding time remaining (Statistics)
            Dim FilePos As Double 'Variable holding the file position NOT IN SEGMENT, IN CHARACTER LENGTH VALUE
            Dim Speed As String 'Variable holding the speed per second in Mb (Statistics)
            TimeRemain = GetTimeRemaining(FileL, Chunk, st, i) 'Get time remaining value (Statistics)
            PercentX = GetPercent(i, segment) 'Get percent value (Statistics)
            Speed = GetSpeed(st, Chunk, i)
            FilePos = i * Chunk 'Number of segment * Characters in a segment
            If i * Chunk > FileL Then
                FilePos = FileL ' Just so it cant go over the file length
            End If
            'This line just show stats in the form caption
            MainFrm.Caption = "SINE256 Visualisation And Sample Console... Decrypting Segment (" & i & "/" & segment & ") (" & FilePos & "/" & FileL & ") (" & PercentX & "%) Time Remaining : " & TimeRemain & "s (Speed : " & Speed & "/s)"
            Dim Bnd As Double 'Buffer calculation (Make sure it doesnt use the full buffer on the last segment which is very rare is the same size as the others)
            If i = segment Then
                Bnd = FileL - (Chunk * (i - 1)) 'Its the last segment so let calculate its real size
            Else
                Bnd = Chunk 'Full buffer
            End If
            ReDim xVal(0 To (Bnd - 1)) 'Resize the Byte array variable
            DoEvents
            If Not EOF(Freenum) Then 'If end of file isnt reached(in case of very exceptionnal error)
                Get #Freenum, , xVal() 'Get data from file with the byte array size which is Bnd
                DoEvents
                cl.SINE256Decrypt xVal(), Res(), password, PatternLen 'Encrypt/Decrypt the current Segment
                DoEvents
                Put #Freenum2, , Res() 'Save Result to the save path
                DoEvents
            End If
        Next ' Loop through i

        Close #Freenum 'Close file (It is now usable)
        Close #Freenum2 'Close file (It is now usable)

         MsgBox "Finished! File length was " & FileL & " characters. Encrypted/Decrypted at " & Round((1000 / (GetTickCount - st)) * FileL) & "Chrs per second", vbInformation

    End If
xStop:
Encrypt.Enabled = True 'Make button possible to use
Decrypt.Enabled = True 'Make button possible to use
EncryptFile.Enabled = True 'Make button possible to use
DecryptFile.Enabled = True 'Make button possible to use
End Sub

Private Sub EncryptFile_Click()
Encrypt.Enabled = False 'Make button impossible to use
Decrypt.Enabled = False 'Make button impossible to use
EncryptFile.Enabled = False 'Make button impossible to use
DecryptFile.Enabled = False 'Make button impossible to use
DoEvents
    'Master variables initializing
    Dim Path As String 'Var holding path of file to encrypt or decrypt
    Dim SavePath As String 'Var holding the save path
    Dim st As Long 'Var holding value of GetTickCount when action start (Statistics)
    Dim Crypt() As Byte 'Variable getting file data (This var get resized and get different size of the file for optimization)
    Dim Res() As Byte 'This is same as Crypt() but its the datas once encrypted (Array is gonna be EXACT SAME SIZE)
    'Show Unused in stats frame
    StatsLbl(0).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(1).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(2).Caption = "Unused" 'Show that those kind of stats arent used with files
    StatsLbl(3).Caption = "Unused" 'Show that those kind of stats arent used with files
    '###Get file to encrypt###
    Dlg.DialogTitle = "Select File to Encrypt/Decrypt"
    Dlg.ShowOpen
    Path = Dlg.Filename
    '#########################
    '####Get path to save#####
    Dlg.DialogTitle = "Enter path to save file"
    Dlg.ShowOpen
    SavePath = Dlg.Filename
    '#########################
    If Not SavePath = "" And Not Path = "" Then
    
        Dim xVal() As Byte 'Buffer variable
        Dim segment As Double 'Number of segment Fix(Size / buffer) + 1
        Dim FileL As Double 'Length of the file
        FileL = LengthOfFile(Path) 'Set length of file
        segment = Fix(FileL / Chunk) + 1 'Set Number of segment
        Dim i As Double 'Variable used to loop (also the position in the segment)
        st = GetTickCount ' sT is the starting gettickcount value . Used for statistics
        Dim Freenum As Integer 'File number of the file to encrypt/decrypt
        Freenum = FreeFile 'Initialize file number
        Dim Freenum2 As Integer 'File number of the path to save
        Dim Str As String
        Open Path For Binary Access Read As #Freenum 'Open the first path before giving file number to save path (otherwise you will get same value for both)
        Freenum2 = FreeFile 'Now you can give him a number
        Open SavePath For Binary Access Write As #Freenum2 'Initialize the save path
        For i = 1 To segment 'Loop in the file buffers (Segment)
            Dim PercentX As Double 'Variable holding percent value (Statistics)
            Dim TimeRemain As Double 'Variable holding time remaining (Statistics)
            Dim FilePos As Double 'Variable holding the file position NOT IN SEGMENT, IN CHARACTER LENGTH VALUE
            Dim Speed As String 'Variable holding the speed per second in Mb (Statistics)
            TimeRemain = GetTimeRemaining(FileL, Chunk, st, i) 'Get time remaining value (Statistics)
            PercentX = GetPercent(i, segment) 'Get percent value (Statistics)
            Speed = GetSpeed(st, Chunk, i)
            FilePos = i * Chunk 'Number of segment * Characters in a segment
            If i * Chunk > FileL Then
                FilePos = FileL ' Just so it cant go over the file length
            End If
            'This line just show stats in the form caption
            MainFrm.Caption = "SINE256 Visualisation And Sample Console... Crypting Segment (" & i & "/" & segment & ") (" & FilePos & "/" & FileL & ") (" & PercentX & "%) Time Remaining : " & TimeRemain & "s (Speed : " & Speed & "/s)"
            Dim Bnd As Double 'Buffer calculation (Make sure it doesnt use the full buffer on the last segment which is very rare is the same size as the others)
            If i = segment Then
                Bnd = FileL - (Chunk * (i - 1)) 'Its the last segment so let calculate its real size
            Else
                Bnd = Chunk 'Full buffer
            End If
            ReDim xVal(0 To (Bnd - 1)) 'Resize the Byte array variable
            DoEvents
            If Not EOF(Freenum) Then 'If end of file isnt reached(in case of very exceptionnal error)
                Get #Freenum, , xVal() 'Get data from file with the byte array size which is Bnd
                DoEvents
                cl.SINE256Encrypt xVal(), Res(), password, PatternLen 'Encrypt/Decrypt the current Segment
                DoEvents
                Put #Freenum2, , Res() 'Save Result to the save path
                DoEvents
            End If
        Next ' Loop through i

        Close #Freenum 'Close file (It is now usable)
        Close #Freenum2 'Close file (It is now usable)

         MsgBox "Finished! File length was " & FileL & " characters. Encrypted/Decrypted at " & Round((1000 / (GetTickCount - st)) * FileL) & "Chrs per second", vbInformation

    End If
xStop:
Encrypt.Enabled = True 'Make button possible to use
Decrypt.Enabled = True 'Make button possible to use
EncryptFile.Enabled = True 'Make button possible to use
DecryptFile.Enabled = True 'Make button possible to use
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



'If you wanna study this encryption ignore the module and everything in here but the encrypt and decrypt button
Private Sub RefreshBaseGraph_Click()
    Dim Xb() As Double
    'This draw the algorithm in a graph (This is the base stuff, nothing is mixed yet)
    Xb() = cl.MakeGraph(decal, start, rangE, PatternLen) 'Initialize the graphic in the class and in the graphic (NEED TO BE INITIALIZED AT LEAST ONCE FOR IT TO WORK)
    ChartIt2 AlgoChart, Xb() 'Refresh the top graph
End Sub

Private Sub SettingCmd_Click()
    SettingFrm.Show
    SettingFrm.SetFocus
End Sub

Private Sub SwapCmd_Click()
    t1.Text = t2.Text 'Simply move text to another text box
    t2.Text = ""
End Sub


Private Sub Decrypt_Click() 'No need to explain , its same as encrypt sub but with the decrypt function
Encrypt.Enabled = False
Decrypt.Enabled = False
EncryptFile.Enabled = False
DecryptFile.Enabled = False
DoEvents
Dim st As Long
If Not t1.Text = "" Then
    Dim Dcrypt() As Byte
    Dim Res() As Byte
    st = GetTickCount
    Dcrypt() = StrConv(t1.Text, vbFromUnicode)
    StatsLbl(0).Caption = GetTickCount - st & "ms"
    st = GetTickCount
    cl.SINE256Decrypt Dcrypt(), Res(), password, PatternLen
    StatsLbl(1).Caption = GetTickCount - st & "ms"
    Dim R As String
    st = GetTickCount
    R = StrConv(Res, vbUnicode)
    StatsLbl(2).Caption = GetTickCount - st & "ms"
    st = GetTickCount
    t2.Text = R
    t1.Text = ""
    StatsLbl(3).Caption = GetTickCount - st & "ms"
End If
Encrypt.Enabled = True
Decrypt.Enabled = True
EncryptFile.Enabled = True
DecryptFile.Enabled = True
End Sub

Private Sub Encrypt_Click()
Encrypt.Enabled = False
Decrypt.Enabled = False
EncryptFile.Enabled = False
DecryptFile.Enabled = False
DoEvents
Dim st As Long
If Not t1.Text = "" Then
    Dim Graph As Boolean
    Select Case MsgBox("Do you want to draw the datas in the graph ? Careful, very long strings can freeze the program. Only use for string under 5000 character for safety.", vbYesNo)
    Case 6 'He clicked YES
        Graph = True
    Case 7 'He clicked No
        Graph = False
    End Select
    Dim Crypt() As Byte 'Variable holding datas to encrypt (need to be in bytes)
    Dim Res() As Byte 'Variable holding the encrypted Data (need to be in bytes)
    st = GetTickCount
    Crypt = StrConv(t1.Text, vbFromUnicode) 'Convert the text in the text box into Bytes array (Its fast!)
    StatsLbl(0).Caption = GetTickCount - st & "ms"
    st = GetTickCount
    cl.SINE256Encrypt Crypt(), Res(), password, PatternLen 'Let's my module convert the stuff
    StatsLbl(1).Caption = GetTickCount - st & "ms"
    st = GetTickCount
    Dim Dat As String
    Dat = StrConv(Res, vbUnicode)
    StatsLbl(2).Caption = GetTickCount - st & "ms"
    st = GetTickCount
    DoEvents
    If Graph = True Then
        ChartIt TextChart, Crypt 'Draw the stuff in the graph
        ChartIt CryptChart, Res 'Draw the stuff in the graph
    End If
    t2.Text = Dat 'Show string
    t1.Text = ""
    StatsLbl(3).Caption = GetTickCount - st & "ms"
End If
Encrypt.Enabled = True
Decrypt.Enabled = True
EncryptFile.Enabled = True
DecryptFile.Enabled = True
End Sub




Private Sub Form_Load()

    Dim z(0 To 1) As Long 'Used to clear the graph with the annoying starting datas
    AlgoChart.ChartData = z 'Clear graph3
    TextChart.ChartData = z 'Clear graph2
    CryptChart.ChartData = z 'Clear graph1
    PatternLen = 50000 'Give patternLen as starting value
    decal = 1 'Give the decal a starting value
    start = 128 'Give the start a starting value
    rangE = 128 'Give the range a starting value
    Chunk = 1048576 'Give buffer size a starting value
        
        ChartIt2 AlgoChart, cl.MakeGraph(decal, start, rangE, PatternLen)
End Sub








Private Sub t1_Change()
    ChrLenLbl.Caption = Len(t1.Text)
End Sub

Private Sub t2_Change()
    ResLenLbl.Caption = Len(t2.Text)
End Sub
