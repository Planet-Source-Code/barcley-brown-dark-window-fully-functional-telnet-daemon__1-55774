VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Black Window 0.1 BETA"
   ClientHeight    =   1725
   ClientLeft      =   11430
   ClientTop       =   1425
   ClientWidth     =   1320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   1320
   Begin MSComCtl2.Animation Anim 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1720
      _Version        =   393216
      Center          =   -1  'True
      BackColor       =   0
      FullWidth       =   73
      FullHeight      =   65
   End
   Begin MSWinsockLib.Winsock Accept 
      Index           =   0
      Left            =   240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Listen 
      Left            =   720
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Active"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserInfo(1 To 25) As String
Dim SocketStatus(1 To 25) As Integer
Dim ReceivedInfo(1 To 25) As String
Dim WordNo As Integer
Dim theWords(0 To 2) As String

Sub SendtoAll(stuff As String)
For a = 1 To 25

If SocketStatus(a) = 1 And UserInfo(a) <> "" Then
Accept(a).SendData stuff
DoEvents

End If
        
        
Next a
End Sub

Private Sub Accept_Close(Index As Integer)
SocketStatus(Index) = 0
UserInfo(Index) = ""
Accept(Index).Close
DoEvents
End Sub

Private Sub Accept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim tmp As String
Dim tmp2 As String

Accept(Index).GetData tmp
DoEvents
Accept(Index).SendData tmp ' ECHOS WHAT IS WRITTEN
DoEvents

If tmp = vbCrLf Or tmp = Chr$(13) Then

If UserInfo(Index) = "" Then

z = 10

If Len(ReceivedInfo(Index)) < z Then z = Len(ReceivedInfo(Index))

tmp2 = Trim(Mid$(ReceivedInfo(Index), 1, z))
tmp2 = Replace(tmp2, " ", "")

     
    tmp3$ = Trim(ReceivedInfo(Index))
    
    For y = 1 To 25
    If UCase$(tmp3$) = UCase$(UserInfo(y)) Then
    Accept(Index).SendData vbCrLf + "This name is already in use." + vbCrLf + "Please enter another: " + vbCrLf
    DoEvents
    ReceivedInfo(Index) = ""
    GoTo skippy
    End If
    Next y

    Call SendtoAll(vbCrLf + "Screen Name: " + tmp2 + " has joined the conversation" + vbCrLf)
    
    UserInfo(Index) = tmp2
    
    Accept(Index).SendData vbCrLf + "Welcome to the conversation." + vbCrLf + "Your screen name is: " + tmp2 + vbCrLf
    DoEvents
    
    Call UserLIST(Index)
     
    
ReceivedInfo(Index) = ""
        
        Else

'==========EXTRA CHAT COMMANDS. USERLISTING, DESCRIPTIONS, ETC, ETC, ET============
'==========COULD ALSO INCLUDE SERVER COMMANDS FOR DATA TRANSFER AND THE LIKE=======


Select Case UCase$(ReceivedInfo(Index))

Case "/USERS"

Call UserLIST(Index)

Case "/QUIT"

Call QUITCHAT(Index)


Case Else

If Mid$(Trim(UCase$(ReceivedInfo(Index))), 1, 8) = "/PRIVATE" Then Call PrivateMessage(Index, ReceivedInfo(Index)): GoTo skippy


Call SendtoAll(vbCrLf + UserInfo(Index) + ": " + ReceivedInfo(Index) + vbCrLf)

skippy:

End Select



'==================================================================================


ReceivedInfo(Index) = ""

End If

Else

ReceivedInfo(Index) = ReceivedInfo(Index) + tmp

End If


End Sub

Sub PrivateMessage(ff As Integer, info As String)

Call WordCount(info)

If WordNo < 2 Then Accept(ff).SendData vbCrLf + "Syntax Error: /PRIVATE <SCREEN NAME> <MESSAGE>" + vbCrLf: DoEvents: GoTo skip

For i = 1 To 25
If UCase$(theWords(1)) = UCase$(UserInfo(i)) Then
Accept(i).SendData vbCrLf + "[Prv.Msg] " + UserInfo(i) + ": " + theWords(2) + vbCrLf
DoEvents
GoTo skip
Else
End If
Next i
Accept(ff).SendData vbCrLf + "Data Error: User specified does not exist" + vbCrLf
DoEvents

skip:
End Sub
Private Sub Accept_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

SocketStatus(Index) = 0
UserInfo(Index) = ""
Accept(Index).Close
DoEvents
End Sub

Sub QUITCHAT(x As Integer)

Accept(x).SendData vbCrLf + "Good-bye " + UserInfo(x) + vbCrLf
DoEvents
SocketStatus(x) = 0
SendtoAll (vbCrLf + "Screen name: " + UserInfo(x) + " has left." + vbCrLf)
Accept(x).Close
DoEvents
End Sub

Sub UserLIST(x As Integer)
Accept(x).SendData vbCrLf + "Participating characters:" + vbCrLf
    DoEvents
    
    For a = 1 To 25

        If SocketStatus(a) = 1 Then
        Accept(x).SendData UserInfo(a) + vbCrLf
        DoEvents
        End If
    Next a
End Sub

Private Sub Form_Load()
Dim Port As Integer
'===COMMAND LINE PORTION==============================
'DarkWindow.exe <Port>
If Command$ = "" Then Port = 33: GoTo 20
Call WordCount(Command$)
If WordNo > 0 And Int(theWords(0)) > 0 Then Port = Int(theWords(0)) Else Port = 33
20

Anim.Open CurDir$ + "\smile.avi" 'Change this in development environment or it can't find file
Anim.AutoPlay = True

'========================================================

'==LOAD LISTENING WINSOCK and load ACCEPTING ARRAY=======
For a = 1 To 25
Load Accept(a)
Next a

Label2.Caption = "PORT:" + Str$(Port)
Listen.LocalPort = Port
Listen.Listen
'========================================================



End Sub

Private Sub Form_Unload(Cancel As Integer)
Close
End Sub

Private Sub Listen_ConnectionRequest(ByVal requestID As Long)
Dim UseSocket As Integer

For a = 1 To 25
If SocketStatus(a) = 0 Then UseSocket = a: SocketStatus(a) = 1: GoTo Accepting
Next a
Listen.Close
DoEvents
Listen.Listen
DoEvents
GoTo 30

Accepting:

Accept(UseSocket).Accept requestID
DoEvents
Accept(UseSocket).SendData vbCrLf + "Welcome to Black Window v0.1 BETA" + vbCrLf
DoEvents
Accept(UseSocket).SendData "Emulation supported: ASCII" + vbCrLf
DoEvents
Accept(UseSocket).SendData "Please enter your desired nickname" + vbCrLf + "(No more than 10 characters): "
DoEvents
30

End Sub

Sub WordCount(text As String)

Dim count As Integer
Dim keepsafe(0 To 2) As String
count = 0
WordNo = 0
spacecount = 0
For a = 0 To 2
theWords(a) = ""
Next a
If Trim(text) = "" Then GoTo 10


message = Trim(text)

For a = 1 To Len(message)

If Mid$(message, a, 1) = " " Then spacecount = spacecount + 1: GoTo SkipALL
If count = 2 Then theWords(2) = Mid$(message, a - 1, (Len(message) - a + 2)): GoTo BreakLoop
If Mid$(message, a, 1) <> " " And spacecount > 0 Then count = count + 1: spacecount = 0

theWords(count) = theWords(count) + Mid$(message, a, 1)

SkipALL:

Next a
BreakLoop:
WordNo = count + 1

10
End Sub

