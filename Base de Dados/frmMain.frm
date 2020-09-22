VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomers 
   BackColor       =   &H00404040&
   Caption         =   " Base de Dados"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Calendario"
      Height          =   735
      Left            =   6360
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "@mail"
      Height          =   735
      Left            =   5400
      Picture         =   "frmMain.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   11
      Left            =   2160
      TabIndex        =   31
      Top             =   4080
      Width           =   2355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   4440
      Picture         =   "frmMain.frx":1C9E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   735
      Left            =   2520
      Picture         =   "frmMain.frx":2968
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5280
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   480
      Left            =   180
      Top             =   4440
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   4210752
      ForeColor       =   -2147483639
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Clientes.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Clientes.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Customers"
      Caption         =   "Contactos em Base de Dados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Combo1 source"
      Connect         =   "Access"
      DatabaseName    =   "Clientes.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3330
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   480
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmdNav 
      BackColor       =   &H8000000A&
      Caption         =   "Adiconar"
      Height          =   735
      Index           =   4
      Left            =   600
      Picture         =   "frmMain.frx":3632
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Dodaj novi zapis"
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Sair"
      Height          =   735
      Index           =   6
      Left            =   3480
      Picture         =   "frmMain.frx":36AC
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Apagar"
      Height          =   735
      Index           =   5
      Left            =   1560
      Picture         =   "frmMain.frx":3740
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   10
      Left            =   5490
      TabIndex        =   22
      Top             =   4080
      Width           =   2250
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   9
      Left            =   2175
      TabIndex        =   20
      Top             =   3675
      Width           =   2355
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   8
      Left            =   5490
      TabIndex        =   18
      Top             =   3675
      Width           =   2250
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PostalCode"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   7
      Left            =   2175
      TabIndex        =   16
      Top             =   3255
      Width           =   2355
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Region"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   6
      Left            =   5730
      TabIndex        =   14
      Top             =   2910
      Width           =   2010
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   5
      Left            =   2175
      TabIndex        =   12
      Top             =   2880
      Width           =   2355
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   4
      Left            =   2175
      TabIndex        =   10
      Top             =   2430
      Width           =   5580
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactTitle"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   5520
      TabIndex        =   8
      Top             =   1995
      Width           =   2235
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactName"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   2175
      TabIndex        =   6
      Top             =   1980
      Width           =   2160
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CompanyName"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   2175
      TabIndex        =   4
      Top             =   1605
      Width           =   5580
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CustomerID"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   2175
      TabIndex        =   2
      Top             =   1215
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Listagem"
      Top             =   480
      Width           =   2940
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label RecNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   6045
      TabIndex        =   27
      Top             =   15
      Width           =   1890
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Procurar Contacto :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   270
      TabIndex        =   26
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   10
      Left            =   4815
      TabIndex        =   21
      Top             =   4080
      Width           =   435
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   9
      Left            =   255
      TabIndex        =   19
      Top             =   3705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Telem:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   8
      Left            =   4800
      TabIndex        =   17
      Top             =   3720
      Width           =   705
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Contribuinte:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   7
      Left            =   255
      TabIndex        =   15
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Localidade:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   13
      Top             =   2940
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Postal:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   5
      Left            =   255
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Morada:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   4
      Left            =   255
      TabIndex        =   9
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Título:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   7
      Top             =   1995
      Width           =   555
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto da Empresa:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   5
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Cliente:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   3
      Top             =   1605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Código do Cliente:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1215
      Width           =   1815
   End
End
Attribute VB_Name = "frmCustomers"
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
Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
RecNo.Caption = Adodc1.Recordset.Bookmark & "/" & Adodc1.Recordset.RecordCount
End Sub
Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Adodc1.Recordset.MoveFirst
Case 1
Adodc1.Recordset.MovePrevious
Case 2
Adodc1.Recordset.MoveNext
Case 3
Adodc1.Recordset.MoveLast
Case 4
Adodc1.Recordset.AddNew
txtFields(0).SetFocus
Case 5
Adodc1.Recordset.Delete
Case 6
End
End Select
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Adodc1.Recordset.Find "CustomerID  = '" & Combo1.Text & "'", , , 0
End Sub
Private Sub Combo1_Click()
Combo1_Change
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1_Change
End Sub
Private Sub Combo1_LostFocus()
Combo1_Change
End Sub

Private Sub Command1_Click()
Unload Me
frmCustomers.Show
End Sub

Private Sub Command2_Click()
Form1.Cliente.Caption = txtFields(0).Text
Dialog.Actual.Caption = txtFields(1).Text
Form1.Designa.Caption = txtFields(1).Text
Form1.contacto.Caption = txtFields(2).Text
Form1.titulo.Caption = txtFields(3).Text
Form1.morada.Caption = txtFields(4).Text
Form1.cp.Caption = txtFields(5).Text
Form1.localidade.Caption = txtFields(6).Text
Form1.nipc.Caption = txtFields(7).Text
Form1.telemovel.Caption = txtFields(8).Text
Form1.telefone.Caption = txtFields(9).Text
Form1.fax.Caption = txtFields(10).Text
'Form1.PrintForm
Form1.Show
Dialog.Show
Unload Me
End Sub

Private Sub Command3_Click()
frmMain.txtTo.Text = txtFields(11).Text
frmMain.Show
Unload Me
End Sub

Private Sub Command4_Click()
frmCal.Show
Unload Me
End Sub

Private Sub Form_Activate()
Data1.Recordset.MoveFirst
  Do Until Data1.Recordset.EOF
   Combo1.AddItem Data1.Recordset.Fields(0).Value, x
   Data1.Recordset.MoveNext
   x = x + 1
 Loop
 Data1.Recordset.MoveFirst
End Sub
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Clientes.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Clientes.mdb"
End Sub

