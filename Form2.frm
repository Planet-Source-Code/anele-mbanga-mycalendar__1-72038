VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MyCalendar"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHighlight 
      Caption         =   "Highlight"
      Height          =   375
      Left            =   5040
      TabIndex        =   25
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtHighlight 
      Height          =   315
      Left            =   3600
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddHoliday 
      Caption         =   "Add Holiday"
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtHoliday 
      Height          =   315
      Left            =   3600
      TabIndex        =   22
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3600
      TabIndex        =   21
      Top             =   4080
      Width           =   1335
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00800080&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   20
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   19
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   18
      Top             =   3720
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFF00&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   14
      Top             =   3360
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0000C000&
      Height          =   255
      Left            =   6480
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   6000
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin VB.ComboBox cboHeaderFormat 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   3600
      List            =   "Form2.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox chkHeaderBold 
      Caption         =   "Header Bold (Toggle)"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   1920
      Width           =   3735
   End
   Begin VB.CheckBox chkMonthBold 
      Caption         =   "Month Bold (Toggle)"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
   End
   Begin VB.CheckBox chkDateBold 
      Caption         =   "Current Date Bold (Toggle)"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CheckBox chkBorder 
      Caption         =   "Border (Toggle)"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox chkSunday 
      Caption         =   "Sunday is a working day"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.CheckBox chkSaturday 
      Caption         =   "Saturday is a working day"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin Project1.MyCalendar MyCalendar 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   5292
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back Color"
      Height          =   195
      Left            =   3600
      TabIndex        =   17
      Top             =   3720
      Width           =   750
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Day Color"
      Height          =   195
      Left            =   3600
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Non Working Day Color"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Format"
      Height          =   195
      Left            =   3600
      TabIndex        =   8
      Top             =   2280
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboHeaderFormat_Click()
    On Error Resume Next
    MyCalendar.HeaderFormat = cboHeaderFormat.ListIndex
    Err.Clear
End Sub

Private Sub chkAppearance_Click()
    On Error Resume Next
    MyCalendar.Appearance = chkBorder.Value
    Err.Clear
End Sub

Private Sub chkBorder_Click()
    On Error Resume Next
    MyCalendar.BorderStyle = chkBorder.Value
    Err.Clear
End Sub

Private Sub chkDateBold_Click()
    On Error Resume Next
    MyCalendar.CurrentDateBold = IIf(chkDateBold.Value = 1, True, False)
    Err.Clear
End Sub

Private Sub chkHeaderBold_Click()
    On Error Resume Next
    MyCalendar.HeaderBold = IIf(chkHeaderBold.Value = 1, True, False)
    Err.Clear
End Sub

Private Sub chkMonthBold_Click()
    On Error Resume Next
    MyCalendar.MonthBold = IIf(chkMonthBold.Value = 1, True, False)
    Err.Clear
End Sub

Private Sub chkSaturday_Click()
    On Error Resume Next
    MyCalendar.SaturdayIsAWorkingDay = IIf(chkSaturday.Value = 1, True, False)
    Err.Clear
End Sub

Private Sub chkSunday_Click()
    On Error Resume Next
    MyCalendar.SundayIsAWorkingDay = IIf(chkSunday.Value = 1, True, False)
    Err.Clear
End Sub

Private Sub cmdAddHoliday_Click()
    On Error Resume Next
    If Len(txtHoliday.Text) = 0 Then
        MsgBox "Please specify the date to add as a holiday.", , "Holiday Error"
        Err.Clear
        Exit Sub
    End If
    
    If IsDate(txtHoliday.Text) = False Then
        MsgBox "The specified value '" & txtHoliday.Text & "' is not a valid date.", , "Holiday Error"
        Err.Clear
        Exit Sub
    End If
    
    MyCalendar.AddHoliday txtHoliday.Text
    Err.Clear
End Sub

Private Sub cmdHighlight_Click()
    On Error Resume Next
    If Len(txtHighlight.Text) = 0 Then
        MsgBox "Please specify the date to highlight.", , "Highlight Error"
        Err.Clear
        Exit Sub
    End If
    
    If IsDate(txtHighlight.Text) = False Then
        MsgBox "The specified value '" & txtHighlight.Text & "' is not a valid date.", , "Highlight Error"
        Err.Clear
        Exit Sub
    End If
    
    MyCalendar.AddHighlight txtHighlight.Text
    Err.Clear
End Sub

Private Sub cmdReset_Click()
    On Error Resume Next
    MyCalendar.Reset
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error Resume Next
    txtHoliday.Text = Format$(DateAdd("d", 2, Now), "dd/mm/yyyy")
    txtHighlight.Text = Format$(DateAdd("d", -5, Now), "dd/mm/yyyy")
    
    MyCalendar.AddHighlight Format$(DateAdd("d", -3, Now), "dd/mm/yyyy")
    MyCalendar.AddHighlight Format$(DateAdd("d", -1, Now), "dd/mm/yyyy")
    MyCalendar.AddHoliday Format$(DateAdd("d", 3, Now), "dd/mm/yyyy")
    Err.Clear
End Sub

Private Sub MyCalendar_DayClick()
    On Error Resume Next
    Caption = "MyCalendar: " & MyCalendar.CurrentDate
    Err.Clear
End Sub

Private Sub Picture1_Click()
    On Error Resume Next
    MyCalendar.NonWorkingBackColor = Picture1.BackColor
    Err.Clear
End Sub

Private Sub Picture2_Click()
    On Error Resume Next
    MyCalendar.NonWorkingBackColor = Picture2.BackColor
    Err.Clear
End Sub

Private Sub Picture3_Click()
    On Error Resume Next
    MyCalendar.NonWorkingBackColor = Picture3.BackColor
    Err.Clear
End Sub

Private Sub Picture4_Click()
    On Error Resume Next
    MyCalendar.CurrentDateBackColor = Picture4.BackColor
    Err.Clear
End Sub

Private Sub Picture5_Click()
    On Error Resume Next
    MyCalendar.CurrentDateBackColor = Picture5.BackColor
    
    Err.Clear
End Sub

Private Sub Picture6_Click()
    On Error Resume Next
    MyCalendar.CurrentDateBackColor = Picture6.BackColor
    
    Err.Clear
End Sub

Private Sub Picture7_Click()
    On Error Resume Next
    MyCalendar.BackColor = Picture7.BackColor
    Err.Clear
End Sub

Private Sub Picture8_Click()
    On Error Resume Next
    MyCalendar.BackColor = Picture8.BackColor
    
    Err.Clear
End Sub

Private Sub Picture9_Click()
    On Error Resume Next
    MyCalendar.BackColor = Picture9.BackColor
    
    Err.Clear
End Sub
