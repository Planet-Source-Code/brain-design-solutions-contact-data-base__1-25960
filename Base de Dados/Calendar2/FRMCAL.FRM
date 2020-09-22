VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCal 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendário"
   ClientHeight    =   4440
   ClientLeft      =   -60
   ClientTop       =   330
   ClientWidth     =   3240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMCAL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   -120
      Width           =   3255
      Begin VB.ComboBox cmbMonth 
         Height          =   330
         IntegralHeight  =   0   'False
         Left            =   110
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Os Meses a visualizar"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtYear 
         Height          =   330
         Left            =   2050
         MaxLength       =   4
         TabIndex        =   8
         ToolTipText     =   "Os Anos a visualizar"
         Top             =   360
         Width           =   855
      End
      Begin VB.PictureBox picTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   110
         ScaleHeight     =   345
         ScaleWidth      =   2985
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   3015
      End
      Begin VB.PictureBox picMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   110
         ScaleHeight     =   2025
         ScaleWidth      =   2985
         TabIndex        =   6
         Top             =   1200
         Width           =   3015
      End
      Begin MSComCtl2.UpDown updwnYear 
         Height          =   330
         Left            =   2895
         TabIndex        =   10
         ToolTipText     =   "Increase/Decrease Year of Appointments to View."
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   100
         Max             =   9998
         Min             =   100
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Calendário Base de Dados"
         ForeColor       =   &H8000000E&
         Height          =   3255
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Transparent
         Height          =   2515
         Left            =   80
         Top             =   680
         Width           =   3100
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000A&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   3255
      Begin Threed.SSCommand SSCommand3 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "visualizar Relógio."
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Relógio"
         ForeColor       =   4210752
         BevelWidth      =   4
      End
      Begin Threed.SSCommand SSCommand2 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Visualizar tabela de Clientes."
         Top             =   720
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Ver &Clientes"
         ForeColor       =   4210752
         BevelWidth      =   4
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "Fechar Calendário."
         Top             =   720
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "&Fechar"
         ForeColor       =   4210752
         BevelWidth      =   4
      End
      Begin Threed.SSCommand btnSeeToday 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Ver o dia actual."
         Top             =   240
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   78
         Caption         =   "Ver dia actual"
         ForeColor       =   4210752
         BevelWidth      =   4
      End
   End
End
Attribute VB_Name = "frmCal"
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

Public DateNum As Date
Public DateSelectedFlag As Boolean

Public YearNum As Long
Public MonthNum As Long
Public DayNum As Long
Public MarkedDay As Long
Public FirstTime As Long
Public MouseX As Single
Public MouseY As Single
Public StartDate As Double
Public Function GetDayOfWeek(ByVal DateToGet As String) As String
  
  Select Case Weekday(CDate(DateToGet), vbSunday)
        
           Case vbSunday
                        GetDayOfWeek = "Domingo"
                        
           Case vbMonday
                        GetDayOfWeek = "Segunda Feira"
          
          Case vbTuesday
                        GetDayOfWeek = "Terça Feira"
        
        Case vbWednesday
                        GetDayOfWeek = "Quarta Feira"
         
         Case vbThursday
                        GetDayOfWeek = "Quinta Feira"
           
           Case vbFriday
                        GetDayOfWeek = "Sexta Feira"
         
         Case vbSaturday
                        GetDayOfWeek = "Sábado"
  
  End Select
  
End Function
Private Sub btnSeeToday_Click()
  
  ' Reset selected date to today '
  DateNum = Now
  
  ' Redraw calendar '
  picTitle.Cls
  picMonth.Cls
  
  Call ResetCalendar
  Call ReDrawCalendar

End Sub
Private Sub Form_Activate()

  Call ResetCalendar
    
End Sub
Private Sub cmbMonth_Click()
  
    ' Get new month number '
    MonthNum = cmbMonth.ListIndex + 1
  
    ' Redraw calendar and mark the first day '
    If IsDate(MonthNum & "/" & DayNum & "/" & YearNum) = False Then
        
      Do
        DayNum = DayNum - 1
         
        If DayNum < 1 Then
            
          DayNum = 1
            
        End If
          
      Loop While IsDate(MonthNum & "/" & DayNum & "/" & YearNum) = False
      
    End If
      
    Call ReDrawCalendar
  
End Sub

Private Sub picMonth_Click()
  
  Dim x As Integer
  Dim y As Integer
  Dim DayBox As Integer
  Dim LastDay As Integer
  Dim FirstDay As Integer
  Dim Serial As Double
  Dim NewDay As Integer
    
  ' Get current location of mouse '
  x = Int(MouseX) + 1
  y = Int(MouseY) + 1
  
  ' Calculate which day-number box mouse is on '
  DayBox = x + (y - 1) * 7
  
  ' Get serial date number for first of month '
  Serial = DateSerial(YearNum, MonthNum, 1)
  
  ' Get last serial date number for month '
  LastDay = DateSerial(YearNum, MonthNum + 1, 1) - Serial
  
  ' Find day of week for first day '
  FirstDay = Weekday(Serial)
  
  ' Determine day number selected '
    NewDay = DayBox - FirstDay + 1
  
  ' Handle selection of blank box before first day '
  If DayBox < FirstDay Then
    
    Beep
    Exit Sub
  
  End If
  
  ' Handle selection of blank box after end of month '
  If DayBox - FirstDay + 1 > LastDay Then
    
    Beep
    Exit Sub
  
  End If
  
  ' Continued if selection passed tests; new day selected '
  DayNum = NewDay
  
  ' Re-mark selected day '
  MarkDay
    
  On Error Resume Next
  picMonth.SetFocus
  On Error GoTo 0
  
End Sub

Private Sub picMonth_GotFocus()
  
    Shape1.BackStyle = 1
    Shape1.BorderStyle = 1
 
End Sub

Private Sub picMonth_KeyDown(KeyCode As Integer, Shift As Integer)
      
  Dim NewDate As Date
  Dim ChangedDate As Boolean
  
  Select Case KeyCode
        
         Case vbKeyLeft
                       NewDate = DateAdd("d", -1, CDate(MonthNum & "/" & DayNum & "/" & YearNum))
                       MonthNum = DatePart("m", NewDate, vbSunday)
                       DayNum = DatePart("d", NewDate, vbSunday)
                       YearNum = DatePart("yyyy", NewDate, vbSunday)
                       ChangedDate = True
                       
        Case vbKeyRight
                       NewDate = DateAdd("d", 1, CDate(MonthNum & "/" & DayNum & "/" & YearNum))
                       MonthNum = DatePart("m", NewDate, vbSunday)
                       DayNum = DatePart("d", NewDate, vbSunday)
                       YearNum = DatePart("yyyy", NewDate, vbSunday)
                       ChangedDate = True
         
         Case vbKeyDown
                       NewDate = DateAdd("d", 7, CDate(MonthNum & "/" & DayNum & "/" & YearNum))
                       MonthNum = DatePart("m", NewDate, vbSunday)
                       DayNum = DatePart("d", NewDate, vbSunday)
                       YearNum = DatePart("yyyy", NewDate, vbSunday)
                       ChangedDate = True
           
           Case vbKeyUp
                       NewDate = DateAdd("d", -7, CDate(MonthNum & "/" & DayNum & "/" & YearNum))
                       MonthNum = DatePart("m", NewDate, vbSunday)
                       DayNum = DatePart("d", NewDate, vbSunday)
                       YearNum = DatePart("yyyy", NewDate, vbSunday)
                       ChangedDate = True
        
  End Select
  
  If ChangedDate = True Then
    
    'gDontChangeDateAtThisTime = True
    
    KeyCode = 0
    If MonthNum - 1 <> frmCal.cmbMonth.ListIndex Then
                         
      frmCal.cmbMonth.ListIndex = MonthNum - 1
                         
    End If
                       
    If YearNum <> frmCal.txtYear.Text Then
                         
      frmCal.txtYear.Text = CStr(YearNum)
                         
    End If
  
    Call ReDrawCalendar
    
    'gDontChangeDateAtThisTime = False
  
  End If
  picMonth.SetFocus
  
End Sub

Private Sub picMonth_LostFocus()
  
  Shape1.BackStyle = 0
  Shape1.BorderStyle = 0

End Sub
Private Sub picMonth_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  
  ' Keep track of mouse location when on calendar '
  MouseX = x
  MouseY = y
  
End Sub

Private Sub SSCommand1_Click()
frmCustomers.Show
Unload Me
End Sub

Private Sub SSCommand2_Click()
frmCustomers.Show
End Sub

Private Sub SSCommand3_Click()
Relogio.Show
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
  
    If KeyAscii < 48 And KeyAscii > 57 Then
    
      If KeyAscii = 13 Then
      
        If Val(txtYear.Text) > 100 And Val(txtYear.Text) < 9998 Then
        
          updwnYear.Value = Val(txtYear.Text)
      
        Else
        
          Call MsgBox("Invalid Year!  Must be 100-9998", vbCritical + vbOKOnly, "Invalid Date.....")
        
        End If
      
      End If
    
      KeyAscii = 0
    
    End If
  
End Sub
Private Sub txtYear_LostFocus()
  
    If Val(txtYear.Text) > 100 And Val(txtYear.Text) < 9998 Then
        
      If Val(txtYear.Text) <> updwnYear.Value Then
      
        updwnYear.Value = Val(txtYear.Text)
      
      End If
    
    Else
        
      Call MsgBox("Invalid Year!  Must be 100-9998", vbCritical + vbOKOnly, "Invalid Date.....")
        
    End If

End Sub
Private Sub updwnYear_Change()
  
  ' Get new year number '
  YearNum = frmCal.updwnYear.Value
  
  ' Redraw calendar and mark the first day '
  If FirstTime = False Then
    
    frmCal.txtYear.Text = Trim(CStr(YearNum))
    
    DayNum = Day(Now)
    Call ReDrawCalendar
    
  End If

End Sub
Private Sub picMonth_Paint()
  
  Call ReDrawCalendar
  
End Sub
Private Sub Form_Load()
  
  frmCal.cmbMonth.AddItem "Janeiro"
  frmCal.cmbMonth.AddItem "Fevereiro"
  frmCal.cmbMonth.AddItem "Março"
  frmCal.cmbMonth.AddItem "Abril"
  frmCal.cmbMonth.AddItem "Maio"
  frmCal.cmbMonth.AddItem "Junho"
  frmCal.cmbMonth.AddItem "Julho"
  frmCal.cmbMonth.AddItem "Agosto"
  frmCal.cmbMonth.AddItem "Setembro"
  frmCal.cmbMonth.AddItem "Outubro"
  frmCal.cmbMonth.AddItem "Novembro"
  frmCal.cmbMonth.AddItem "Dezembro"
  
End Sub
Public Sub MarkDay()
  
  Dim TheDay As Integer
  Dim i As Integer
  Dim Serial As Double
  Dim DayBox As Integer
  Dim x As Integer
  Dim y As Integer
  Dim x1 As Single
  Dim y1 As Single
  Dim x2 As Single
  Dim y2 As Single
  
  ' Record day number '
  TheDay = DayNum
  
  ' Erase previous mark, then mark current day '
  
  For i = 1 To 2
     
     ' Calculate box number for day '
     Serial = DateSerial(YearNum, MonthNum, 1)
     DayBox = Weekday(Serial) + DayNum - 1
     
     ' Calculate location of box number '
     x = ((DayBox - 1) Mod 7) + 1
     y = ((DayBox - 1) \ 7) + 1
     
     ' Get first corner location of box '
     x1 = (x% - 1) * frmCal.picMonth.ScaleWidth / 7
     y1 = (y% - 1) * frmCal.picMonth.ScaleHeight / 6
     
     ' Get second corner location of box '
     x2 = x1 + frmCal.picMonth.ScaleWidth / 7
     y2 = y1 + frmCal.picMonth.ScaleHeight / 6
     
     ' XOR box pixels '
     frmCal.picMonth.DrawMode = 7
     frmCal.picMonth.Line (x1, y1)-(x2, y2), QBColor(15), BF
     
     ' Quit if no previously marked day '
     If MarkedDay = 0 Then
       
       Exit For
       
     ' Prepare to mark currently selected day '
     Else
       
       DayNum = MarkedDay
     
     End If
  
  Next i
  
  ' Reset day number '
  DayNum = TheDay
  
  ' Record marked day for next trip through here '
  MarkedDay = DayNum
  
End Sub
Public Sub PutNum(Square As Integer, Num As Integer)
  
  Dim n As String
  Dim x As Integer
  Dim y As Integer
  
  ' Build string of day-number digits '
  n = LTrim$(Str$(Num))
  
  ' Calculate location of box '
  x = ((Square - 1) Mod 7) + 1
  y = (Square - 1) \ 7 + 1
  
  ' Set print position '
  frmCal.picMonth.CurrentX = x - 0.5 - frmCal.picMonth.TextWidth(n$) / 2
  frmCal.picMonth.CurrentY = y - 0.5 - frmCal.picMonth.TextHeight(n$) / 2
  
  ' Display day number '
  frmCal.picMonth.Print n
  
End Sub
Public Sub ReDrawCalendar()
  
  ' Clear out previous stuff '
  frmCal.picMonth.Cls
  frmCal.picTitle.Cls
  
  ' Redraw calendar '
  MarkedDay = 0
  
  ' Display names of days at top '
  Call WeekDayNames
    
  ' Draw lines that form calendar '
  Call DrawLines
  
  ' Fill calendar with day numbers '
  Call FillCal
  
  ' Mark currently selected day '
  Call MarkDay
  
End Sub
Public Sub WeekDayNames()
  
  Dim i As Integer
  Dim D As String
  
  ' Scale for displaying seven day names '
  frmCal.picTitle.Scale (0, 0)-(7, 1)
  
  ' Display each weekday name '
  For i = 0 To 6
     
     ' Get three-letter abbreviation '
      D = Format$(CDbl(i + 1), "ddd")
      
      ' Use two characters if user's font is too wide '
      If frmCal.picTitle.TextWidth("Wed") > 1 Then
        
        D = Left$(D, 2)
      
      End If
      
      ' Display each weekday name '
      frmCal.picTitle.CurrentX = i + 0.5 - frmCal.picTitle.TextWidth(D) / 2
      frmCal.picTitle.CurrentY = 1 - frmCal.picTitle.TextHeight(D)
      frmCal.picTitle.Print D
      
  Next i
  
End Sub
Public Sub ResetCalendar()
  
  ' Default to no date selected '
  DateSelectedFlag = False
  
  ' Record starting date '
  StartDate = DateNum
  
  ' If no starting date, use current date '
  If DateNum = 0 Then
    
    DateNum = Now
    
  End If
    
  ' Keep track of marked day '
  MarkedDay = 0
    
  ' Extract date parts '
  YearNum = Year(DateNum)
  MonthNum = Month(DateNum)
  DayNum = Day(DateNum)
    
  ' Flag that this is first time through '
  FirstTime = True
   
  ' Initialize scrollbars '
  frmCal.updwnYear.Value = YearNum
  frmCal.cmbMonth.ListIndex = MonthNum - 1
  frmCal.txtYear.Text = YearNum
  
  ' Flag first time setting scrollbars is done '
  FirstTime = False

End Sub
Public Sub FillCal()
  
  Dim NumDays As Long
  Dim Serial1 As Double
  Dim Serial2 As Double
  Dim DayOffset As Integer
  Dim Daycounter As Integer
  
  ' Get serial date number for 1st of current month '
  Serial1 = DateSerial(YearNum, MonthNum, 1)
  
  ' Get serial date number for 1st of next month '
  Serial2 = DateSerial(YearNum, MonthNum + 1, 1)
  
  ' Calculate number of days in month '
  NumDays = Serial2 - Serial1
  
  ' Skip over blank days at start of month '
  DayOffset = Weekday(Serial1) - 1
  
  For Daycounter = 1 To NumDays
     
     PutNum Daycounter + DayOffset, Daycounter
     
  Next Daycounter
  
End Sub
Public Sub DrawLines()
  
  Dim x As Integer
  Dim y As Integer
  
  ' Draw lines that separate days '
  frmCal.picMonth.Scale (0, 0)-(7, 6)
  frmCal.picMonth.DrawMode = 13
  
  ' Draw vertical lines '
  For x = 1 To 6
     
     frmCal.picMonth.Line (x, 0)-(x, 6)
     frmCal.picTitle.Line (x, 0)-(x, 6)
     
  Next x
  
  ' Draw horizontal lines '
  For y = 1 To 5
     
     frmCal.picMonth.Line (0, y)-(7, y)
     
  Next y
  
End Sub
