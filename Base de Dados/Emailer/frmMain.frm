VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "@mail"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   DrawStyle       =   6  'Inside Solid
   ForeColor       =   &H8000000E&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand cmdExit 
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   5880
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Fechar"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
   End
   Begin Threed.SSCommand cmdSend 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   5880
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Enviar @mail"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   375
      Left            =   5000
      TabIndex        =   12
      Top             =   1080
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Ver Clientes"
      ForeColor       =   -2147483642
      BevelWidth      =   4
   End
   Begin VB.Timer tmrClearStatus 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   0
      Top             =   6000
   End
   Begin MSComDlg.CommonDialog cdSave 
      Left            =   480
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrLoadDefault 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5640
      Top             =   6120
   End
   Begin VB.Timer tmrLogo1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5160
      Top             =   6120
   End
   Begin VB.Timer tmrLogo 
      Interval        =   500
      Left            =   4680
      Top             =   6120
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   5640
      Top             =   5640
   End
   Begin VB.TextBox txtFrom 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Brain Design Solutions.com"
      ToolTipText     =   "O seu email."
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtSubject 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "O assunto do email."
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtTo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Escreva aqui o destinatário"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4680
      Top             =   5640
   End
   Begin VB.Timer tmrCheckFields 
      Interval        =   1
      Left            =   0
      Top             =   5520
   End
   Begin VB.TextBox txtBody 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      ToolTipText     =   "Escreva aqui a mensagem!"
      Top             =   2040
      Width           =   5955
   End
   Begin VB.TextBox txtServer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "smtp.iol.pt"
      ToolTipText     =   "o servidor de acesso!"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Estado de envio:"
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   6135
      Begin VB.Label lblSendingStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5775
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5160
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Opções:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4960
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblSaveEmail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salvar @mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Base de Dados Clientes 1.0.5
'OFPG - Marketing e Publicidade, Lda - Brain Design Solutions
'http://www.braindesignsolutions.com
'softwaredirector@braindesignsolutions.com
'Agosto 2001
'Made in Portugal (UE)
'-------------------------------------------
Dim Response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim Start As Single, Tmr As Single

Private Sub cmdClear_Click()
txtTo = ""
txtFrom = ""
txtSubject = ""
txtServer = ""
txtBody = ""
cmdSend.Enabled = True
lblSendingStatus.Caption = ""
tmrAnimate.Enabled = False
tmrTimeOut.Enabled = False
txtTo.SetFocus
End Sub

Private Sub cmdExit_Click()
Unload Me
frmCustomers.Show
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0&
End Sub

Private Sub cmdSend_Click()

tmrCheckFields.Enabled = True
cmdSend.Enabled = True
tmrTimeOut.Enabled = True
lblSendingStatus.Caption = "A enviar email..."
tmrAnimate.Enabled = True

SendEmail txtTo, txtFrom, txtSubject, txtBody, txtServer
tmrAnimate.Enabled = False
lblSendingStatus.Caption = " Email enviado! "
Beep
tmrCheckFields.Enabled = True

tmrClearStatus.Enabled = True
tmrTimeOut.Enabled = False
Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H8000000E
End Sub

Private Sub imgLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H8000000E
End Sub


Private Sub lblAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblAbout.FontSize = 10
lblAbout.ForeColor = &HFF&
End Sub

Private Sub lblSaveEmail_Click()

cdSave.Filter = "Text files (*.txt)|*.txt|"

cdSave.FileName = "Email"

cdSave.ShowSave

Open cdSave.FileName For Append As #1
Print #1, "-------------------------"
Print #1, txtTo
Print #1, txtFrom
Print #1, txtSubject
Print #1, txtServer
Print #1,
Print #1, txtBody
Print #1, "-------------------------"
Close #1
End Sub

Private Sub lblSaveEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 10
lblSaveEmail.ForeColor = &HFF&
End Sub

Private Sub SSCommand1_Click()
frmCustomers.Show
End Sub

Private Sub SSCommand2_Click()

tmrCheckFields.Enabled = True
cmdSend.Enabled = True
tmrTimeOut.Enabled = True
lblSendingStatus.Caption = "A enviar email"
tmrAnimate.Enabled = True

SendEmail txtTo, txtFrom, txtSubject, txtBody, txtServer
tmrAnimate.Enabled = False
lblSendingStatus.Caption = " Email enviado! "
Beep
tmrCheckFields.Enabled = True

tmrClearStatus.Enabled = True
tmrTimeOut.Enabled = False
Close
End Sub

Private Sub tmrAnimate_Timer()

lblSendingStatus.Caption = lblSendingStatus.Caption + " ."
End Sub

Private Sub tmrCheckFields_Timer()

If txtFrom <> "" And txtTo <> "" And txtSubject <> "" And txtServer <> "" And txtBody <> "" Then
 cmdSend.Enabled = True
Else
 cmdSend.Enabled = False
End If
End Sub

Private Sub tmrClearStatus_Timer()
lblSendingStatus.Caption = ""
tmrClearStatus.Enabled = False
End Sub


Private Sub tmrLoadDefault_Timer()
tmrLoadDefault.Enabled = False
tmrLogo.Enabled = True
End Sub

Private Sub tmrLogo_Timer()
tmrLogo.Enabled = False
tmrLogo1.Enabled = True
End Sub

Private Sub tmrLogo1_Timer()
tmrLogo1.Enabled = False
tmrLoadDefault.Enabled = True
End Sub

Private Sub tmrTimeOut_Timer()
lblSendingStatus.Caption = "Error"
tmrAnimate.Enabled = False
MsgBox "A ligação ao servidor falhou. Certifique que tudo está preenchido correctamente!", vbCritical, "Connection timed out"
End Sub

Private Sub txtFrom_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &HFFFFFF
End Sub

Private Sub txtServer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &HFFFFFF
End Sub

Private Sub txtSubject_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0&
End Sub

Private Sub txtTo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

lblSaveEmail.FontSize = 8
lblSaveEmail.ForeColor = &H0&
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData Response
End Sub


Sub SendEmail(EmailTo As String, From As String, Subject As String, Body As String, Server As String)
          
Winsock1.LocalPort = 0
    
If Winsock1.State = sckClosed Then
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + txtFrom + vbCrLf
    Second = "rcpt to:" + Chr(32) + txtTo + vbCrLf
    Third = "Date:" + Chr(32) + DateNow + vbCrLf
    Fourth = "From:" + Chr(32) + vbCrLf
    Fifth = "Para:" + Chr(32) + txtTo + vbCrLf
    Sixth = "Assunto:" + Chr(32) + txtSubject + vbCrLf
    Seventh = "Mensagem:" + Chr(32) + txtBody + vbCrLf
    Ninth = "Brain Mail 1.0 - Brain Design Solutions" + vbCrLf
    Eighth = Fourth + Third + Ninth + Fifth + Sixth

    Winsock1.Protocol = sckTCPProtocol
    Winsock1.RemoteHost = txtServer
    Winsock1.RemotePort = 25
    Winsock1.Connect
    
    WaitFor ("220")
    
    lblSendingStatus.Caption = "A Ligar ao Servidor"
    
    Winsock1.SendData ("HELO hotmail.com" + vbCrLf)

    WaitFor ("250")

    lblSendingStatus.Caption = "Ligado ao Servidor"
    
    Winsock1.SendData (first)

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    Start = Timer
    While Len(Response) = 0
        Tmr = Start - Timer
        DoEvents
        If Tmr > 50 Then
            MsgBox "O tempo de espera na ligação ao Servidor SMTP foi excedido.", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(Response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "O Servidor SMTP não respondeu devidamente: " + ResponseCode + " Code recieved: " + Response, 64, MsgTitle
            Exit Sub
        End If
    Wend
Response = ""
End Sub


