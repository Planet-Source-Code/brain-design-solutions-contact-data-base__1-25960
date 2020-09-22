VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3030
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   5025
      Begin VB.CommandButton Command1 
         Caption         =   "&Continuar"
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblCopyright 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Copyright 2001 OFPG - Marketing e Publicidade, Lda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Licença a OFPG - Marketing e Publicidade, Lda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Base de Dados"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   840
         TabIndex        =   3
         Top             =   480
         Width           =   1950
      End
   End
End
Attribute VB_Name = "frmSplash"
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
Option Explicit

Private Sub Command1_Click()
Unload Me
frmCustomers.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub



Private Sub Frame1_Click()
    Unload Me
End Sub

