VERSION 5.00
Begin VB.Form Relogio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relógio"
   ClientHeight    =   285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1560
   Icon            =   "Relogio.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1560
      Top             =   0
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Relogio"
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
Private Sub lblTime_Click()

End Sub

Private Sub Timer1_Timer()
lblTime.Caption = "" & Time & ""
End Sub
