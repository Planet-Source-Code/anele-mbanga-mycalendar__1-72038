VERSION 5.00
Begin VB.UserControl MyCalendar 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ScaleHeight     =   2880
   ScaleWidth      =   3360
   Begin VB.CommandButton cmdNextYear 
      Height          =   315
      Left            =   2640
      Picture         =   "MyCalendar.ctx":0000
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Following year"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdPreviousYear 
      Height          =   315
      Left            =   320
      Picture         =   "MyCalendar.ctx":0596
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Previous year"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox txtMonth 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   720
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton cmdNextDay 
      Height          =   315
      Left            =   3000
      Picture         =   "MyCalendar.ctx":0B2C
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Following month"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cmdPreviousDay 
      Height          =   315
      Left            =   0
      Picture         =   "MyCalendar.ctx":10AE
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Previous month"
      Top             =   0
      Width           =   315
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   42
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":1610
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   41
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":1762
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   40
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":18B4
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   39
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":1A06
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   38
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":1B58
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   37
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":1CAA
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   36
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":1DFC
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   2520
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   35
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":1F4E
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   34
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":20A0
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   33
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":21F2
      MousePointer    =   99  'Custom
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   32
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":2344
      MousePointer    =   99  'Custom
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   31
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":2496
      MousePointer    =   99  'Custom
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   30
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":25E8
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   29
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":273A
      MousePointer    =   99  'Custom
      TabIndex        =   35
      Text            =   "Text1"
      Top             =   2160
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   28
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":288C
      MousePointer    =   99  'Custom
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   27
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":29DE
      MousePointer    =   99  'Custom
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   26
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":2B30
      MousePointer    =   99  'Custom
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   25
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":2C82
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   24
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":2DD4
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   23
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":2F26
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   22
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":3078
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   1800
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   21
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":31CA
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   20
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":331C
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   19
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":346E
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   18
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":35C0
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   17
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":3712
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   16
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":3864
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   15
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":39B6
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   1440
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   14
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":3B08
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   13
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":3C5A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   12
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":3DAC
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   11
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":3EFE
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   10
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":4050
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   9
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":41A2
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   8
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":42F4
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   7
      Left            =   2880
      MouseIcon       =   "MyCalendar.ctx":4446
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   6
      Left            =   2400
      MouseIcon       =   "MyCalendar.ctx":4598
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   5
      Left            =   1920
      MouseIcon       =   "MyCalendar.ctx":46EA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   4
      Left            =   1440
      MouseIcon       =   "MyCalendar.ctx":483C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   3
      Left            =   960
      MouseIcon       =   "MyCalendar.ctx":498E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   2
      Left            =   480
      MouseIcon       =   "MyCalendar.ctx":4AE0
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox DayOf 
      Height          =   315
      Index           =   1
      Left            =   0
      MouseIcon       =   "MyCalendar.ctx":4C32
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      CausesValidation=   0   'False
      Height          =   315
      Index           =   7
      Left            =   2880
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   6
      Left            =   2400
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   5
      Left            =   1920
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   3
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   360
      Width           =   435
   End
   Begin VB.TextBox Header 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Sun"
      Top             =   360
      Width           =   435
   End
End
Attribute VB_Name = "MyCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************************************'
'   Developed by Anele Mbanga
'   anelem@rocketmail.com
'   There is no guarantee of this source code, thus use at own risk
'   However if you use, you can give credit to me
'   To Do...
'   Change the previous year, previous month etc to be graphics, etc
'
'**********************************************************

Public Enum HeaderFormatEnum
    ddd = 0             ' Mon(day)
    dd = 1              ' Mo(nday)
    d = 2               ' M(onday)
End Enum

Public Enum BorderStyleEnum
    None = 0
    Fixed = 1
End Enum

Public Enum AppearanceEnum
    Flat = 0
    ThreeD = 1
End Enum

Private Type Holiday
    HD_ID As String
    HD_Name As String
End Type

Private mvarHeaderFormat As HeaderFormatEnum
Private mvarBorderStyle As BorderStyleEnum
Private mvarHeaderBold As Boolean
Private mvarMonthBold As Boolean
Private mvarAppearance As AppearanceEnum
Private mvarCurrentDate As String
Private mvarCurrentYear As Integer
Private mvarCurrentMonth As Integer
Private mvarCurrentDay As Integer
Private mvarCurrentDateBackColor As OLE_COLOR
Private mvarCurrentDateBold As Boolean
Private mvarSaturdayIsAWorkingDay As Boolean
Private mvarSundayIsAWorkingDay As Boolean
Private mvarBackColor As OLE_COLOR
Private mvarNonWorkingBackColor As OLE_COLOR
Private mvarDateFormat As String
Event DayClick()
Event MonthChange()
Event YearChange()
Private sPos As Integer
Private colHighlights As Collection
Private colWorking As Collection
Private HolidayCount As Long
Private Holidays() As Holiday
Private mvarStartDate As String
Private mvarEndDate As String
Private mvarWorkingDays As Integer

Public Sub AddHighlight(ByVal Highlight As String)
    On Error Resume Next
    ' add a date to be highlighted bold
    Highlight = Format$(Highlight, DateFormat)
    If IsDate(Highlight) = True Then colHighlights.Add Highlight, Highlight
    MarkHighlights
    Err.Clear
End Sub

Public Sub AddWorkingDay(ByVal WorkingDate As String)
    On Error Resume Next
    ' add a date to be highlighted bold
    WorkingDate = Format$(WorkingDate, DateFormat)
    If IsDate(WorkingDate) = True Then colWorking.Add WorkingDate, WorkingDate
    MarkWorkingDays
    Err.Clear
End Sub


Public Sub AddHoliday(ByVal HolidayDate As String, ByVal HolidayName As String)
    On Error Resume Next
    ' adds a specified date to the holidays collection
    HolidayDate = Format$(HolidayDate, DateFormat)
    If IsDate(HolidayDate) = True Then
        HolidayCount = HolidayCount + 1
        ReDim Preserve Holidays(HolidayCount)
        Holidays(HolidayCount).HD_ID = HolidayDate
        Holidays(HolidayCount).HD_Name = HolidayName
    End If
    ' highlight all holidays with the non working days color
    MarkHolidays
    Err.Clear
End Sub

Private Sub MarkHighlights()
    On Error Resume Next
    'mark all specified dates bold
    Dim rsTot As Integer
    Dim rsStr As String
    Dim dPos As Integer
    Dim rsCnt As Integer
    
    ' count all highlights
    rsTot = colHighlights.Count
    For rsCnt = 1 To rsTot
        rsStr = colHighlights(rsCnt)
        If Len(rsStr) > 0 Then
            ' search for the day
            dPos = DateSearch(rsStr)
            If dPos > 0 Then DayOf(dPos).Font.Bold = True
        End If
        DoEvents
        Err.Clear
    Next
    Err.Clear
End Sub

Private Sub MarkHolidays()
    On Error Resume Next
    'mark all holidays with the non working days color
    Dim rsStr As String
    Dim dPos As Integer
    Dim rsCnt As Long
    
    ' load all holidays
    For rsCnt = 1 To HolidayCount
        rsStr = Holidays(rsCnt).HD_ID
        If Len(rsStr) > 0 Then
            ' search for the holiday
            dPos = DateSearch(rsStr)
            If dPos > 0 Then
                DayOf(dPos).BackColor = mvarNonWorkingBackColor
                DayOf(dPos).ToolTipText = Holidays(rsCnt).HD_Name
            End If
        End If
        DoEvents
        Err.Clear
    Next
    Err.Clear
End Sub

Private Sub MarkWorkingDays()
    On Error Resume Next
    'mark all holidays with the non working days color
    Dim rsStr As String
    Dim dPos As Integer
    Dim rsCnt As Long
    
    ' load all working days
    ' these are days that may fall within a weekend which is marked as non working
    For rsCnt = 1 To colWorking.Count
        rsStr = colWorking(rsCnt)
        If Len(rsStr) > 0 Then
            ' search for the working day
            dPos = DateSearch(rsStr)
            If dPos > 0 Then
                DayOf(dPos).BackColor = &H80000005
                DayOf(dPos).ToolTipText = ""
            End If
        End If
        DoEvents
        Err.Clear
    Next
    Err.Clear
End Sub

Public Sub ClearHighlights()
    On Error Resume Next
    ' clears all highlights
    Dim rsTot As Integer
    Dim rsStr As String
    Dim dPos As Integer
    Dim rsCnt As Integer
    
    rsTot = colHighlights.Count
    For rsCnt = 1 To rsTot
        rsStr = colHighlights(rsCnt)
        If Len(rsStr) > 0 Then
            dPos = DateSearch(rsStr)
            If dPos > 0 Then DayOf(dPos).Font.Bold = False
        End If
        DoEvents
        Err.Clear
    Next
    Set colHighlights = New Collection
    Err.Clear
End Sub

Public Sub ClearHolidays()
    On Error Resume Next
    ' clears all holidays
    Dim rsStr As String
    Dim dPos As Integer
    Dim rsCnt As Long
    
    For rsCnt = 1 To HolidayCount
        rsStr = Holidays(rsCnt).HD_ID
        If Len(rsStr) > 0 Then
            dPos = DateSearch(rsStr)
            If dPos > 0 Then
                DayOf(dPos).BackColor = &H80000005
                DayOf(dPos).ToolTipText = ""
            End If
        End If
        DoEvents
        Err.Clear
    Next
    ReDim Holidays(0)
    Err.Clear
End Sub

Private Function DateSearch(ByVal SearchDate As String) As Integer
    On Error Resume Next
    ' returns the position of the day we are looking for
    Dim rsCnt As Integer
    Dim rsStr As String
    
    DateSearch = -1
    For rsCnt = 1 To 42
        rsStr = DayOf(rsCnt).Tag
        If SearchDate = rsStr Then
            DateSearch = rsCnt
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function

Private Sub SetHeader()
    On Error Resume Next
    ' set headers for the calendar
    Dim rsCnt As Integer
    For rsCnt = 1 To 7
        Header(rsCnt).Font.Bold = mvarHeaderBold
        Header(rsCnt).Locked = True
        Err.Clear
    Next
    
    Select Case mvarHeaderFormat
    Case ddd
        Header(1).Text = "Sun"
        Header(2).Text = "Mon"
        Header(3).Text = "Tue"
        Header(4).Text = "Wed"
        Header(5).Text = "Thu"
        Header(6).Text = "Fri"
        Header(7).Text = "Sat"
    Case dd
        Header(1).Text = "Su"
        Header(2).Text = "Mo"
        Header(3).Text = "Tu"
        Header(4).Text = "We"
        Header(5).Text = "Th"
        Header(6).Text = "Fr"
        Header(7).Text = "Sa"
    Case d
        Header(1).Text = "S"
        Header(2).Text = "M"
        Header(3).Text = "T"
        Header(4).Text = "W"
        Header(5).Text = "T"
        Header(6).Text = "F"
        Header(7).Text = "S"
    End Select
    Err.Clear
End Sub

Public Property Get DateFormat() As String
    On Error Resume Next
    ' format for date
    DateFormat = mvarDateFormat
    Err.Clear
End Property

Public Property Let DateFormat(ByVal NewValue As String)
    On Error Resume Next
    mvarDateFormat = NewValue
    mvarCurrentDate = Format$(mvarCurrentDate, mvarDateFormat)
    PropertyChanged "DateFormat"
    Err.Clear
End Property

Public Property Get SaturdayIsAWorkingDay() As Boolean
    On Error Resume Next
    SaturdayIsAWorkingDay = mvarSaturdayIsAWorkingDay
    Err.Clear
End Property

Public Property Let SaturdayIsAWorkingDay(ByVal NewValue As Boolean)
    On Error Resume Next
    Dim rsCnt As Integer
    'stores whether saturday is a working day
    'if not, it will be marked with the non working day color
    mvarSaturdayIsAWorkingDay = NewValue
    For rsCnt = 1 To 42
        Select Case rsCnt
        Case 7, 14, 21, 28, 35, 42
            ' saturday
            DayOf(rsCnt).BackColor = IIf(mvarSaturdayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        Case 1, 8, 15, 22, 29, 36
            ' sunday
            DayOf(rsCnt).BackColor = IIf(mvarSundayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        End Select
        If Len(DayOf(rsCnt)) = 0 Then
            DayOf(rsCnt).BackColor = UserControl.BackColor
            DayOf(rsCnt).Enabled = False
        End If
        Err.Clear
    Next
    
    DayOf(mvarCurrentDay + sPos - 1).BackColor = mvarCurrentDateBackColor
    DayOf(mvarCurrentDay + sPos - 1).Font.Bold = mvarCurrentDateBold
    
    PropertyChanged "SaturdayIsAWorkingDay"
    Err.Clear
End Property

Public Property Get SundayIsAWorkingDay() As Boolean
    On Error Resume Next
    SundayIsAWorkingDay = mvarSundayIsAWorkingDay
    Err.Clear
End Property

Public Property Let SundayIsAWorkingDay(ByVal NewValue As Boolean)
    On Error Resume Next
    Dim rsCnt As Integer
    mvarSundayIsAWorkingDay = NewValue
    'stores whether sunday is a working day
    'if not, it will be marked with the non working day color
    For rsCnt = 1 To 42
        Select Case rsCnt
        Case 7, 14, 21, 28, 35, 42
            ' saturday
            DayOf(rsCnt).BackColor = IIf(mvarSaturdayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        Case 1, 8, 15, 22, 29, 36
            ' sunday
            DayOf(rsCnt).BackColor = IIf(mvarSundayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        End Select
        If Len(DayOf(rsCnt)) = 0 Then
            DayOf(rsCnt).BackColor = UserControl.BackColor
            DayOf(rsCnt).Enabled = False
        End If
        Err.Clear
    Next
    
    DayOf(mvarCurrentDay + sPos - 1).BackColor = mvarCurrentDateBackColor
    DayOf(mvarCurrentDay + sPos - 1).Font.Bold = mvarCurrentDateBold
    PropertyChanged "SundayIsAWorkingDay"
    Err.Clear
End Property


Public Property Get CurrentDateBold() As Boolean
    On Error Resume Next
    CurrentDateBold = mvarCurrentDateBold
    Err.Clear
End Property

Public Property Let CurrentDateBold(ByVal NewValue As Boolean)
    On Error Resume Next
    ' when true,this highlights the current selected date as bold
    mvarCurrentDateBold = NewValue
    DayOf(mvarCurrentDay + sPos - 1).Font.Bold = NewValue
    PropertyChanged "CurrentDateBold"
    DayOf(mvarCurrentDay + sPos - 1).Refresh
    Err.Clear
End Property

Public Property Get HeaderBold() As Boolean
    On Error Resume Next
    HeaderBold = mvarHeaderBold
    Err.Clear
End Property

Public Property Let HeaderBold(ByVal NewValue As Boolean)
    On Error Resume Next
    Dim rsCnt As Integer
    ' when true, marks the headers as bold
    mvarHeaderBold = NewValue
    For rsCnt = 1 To 7
        Header(rsCnt).Font.Bold = mvarHeaderBold
        Err.Clear
    Next
    PropertyChanged "HeaderBold"
    Err.Clear
End Property

Public Property Get MonthBold() As Boolean
    On Error Resume Next
    MonthBold = mvarMonthBold
    Err.Clear
End Property

Public Property Let MonthBold(ByVal NewValue As Boolean)
    On Error Resume Next
    ' when true, marks the month title as bold e.g April 2009
    mvarMonthBold = NewValue
    txtMonth.Font.Bold = NewValue
    PropertyChanged "MonthBold"
    Err.Clear
End Property


Public Property Get HeaderFormat() As HeaderFormatEnum
    On Error Resume Next
    HeaderFormat = mvarHeaderFormat
    Err.Clear
End Property

Public Property Let HeaderFormat(ByVal NewValue As HeaderFormatEnum)
    On Error Resume Next
    ' sets the header format between ddd for Thu (rsday), dd for Th (ursday) and d for T (hursday)
    mvarHeaderFormat = NewValue
    SetHeader
    PropertyChanged "HeaderFormat"
    Err.Clear
End Property

Public Property Get BorderStyle() As BorderStyleEnum
    On Error Resume Next
    BorderStyle = mvarBorderStyle
    Err.Clear
End Property

Public Property Let BorderStyle(ByVal NewValue As BorderStyleEnum)
    On Error Resume Next
    ' sets border of control between flat and threed
    mvarBorderStyle = NewValue
    UserControl.BorderStyle = NewValue
    PropertyChanged "BorderStyle"
    Err.Clear
End Property

Private Sub cmdNextDay_Click()
    On Error Resume Next
    ' the next monthbutton has been clicked, show the next month, same year,same day
    CurrentDate = DateAdd("m", 1, CurrentDate)
    RaiseEvent MonthChange
    Err.Clear
End Sub

Private Sub cmdNextYear_Click()
    On Error Resume Next
    ' the next year has been clicked, show the next year, same month, same day
    CurrentDate = DateAdd("yyyy", 1, CurrentDate)
    RaiseEvent YearChange
    Err.Clear
End Sub

Private Sub cmdPreviousDay_Click()
    On Error Resume Next
    ' previous month has been selected, show
    CurrentDate = DateAdd("m", -1, CurrentDate)
    RaiseEvent MonthChange
    Err.Clear
End Sub

Private Sub cmdPreviousYear_Click()
    On Error Resume Next
    'previous year has been selected, show
    CurrentDate = DateAdd("yyyy", -1, CurrentDate)
    RaiseEvent YearChange
    Err.Clear
End Sub

Private Sub DayOf_Click(Index As Integer)
    On Error Resume Next
    ' a day has been selected as current date
    ' select it
    CurrentDate = DayOf(Index).Tag
    mvarCurrentDate = CurrentDate
    UserControl.Extender.SetFocus
    RaiseEvent DayClick
    Err.Clear
End Sub

Private Sub ClearDays()
    On Error Resume Next
    ' clear calendar days
    Dim rsCnt As Integer
    For rsCnt = 1 To 42
        DayOf(rsCnt).Text = ""
        DayOf(rsCnt).Alignment = 2
        DayOf(rsCnt).Locked = True
        DayOf(rsCnt).BackColor = &H80000005
        DayOf(rsCnt).Tag = ""
        DayOf(rsCnt).ToolTipText = ""
        DayOf(rsCnt).TabStop = False
        DayOf(rsCnt).Font.Bold = False
        DayOf(rsCnt).Enabled = True
        Err.Clear
    Next
    Err.Clear
End Sub

Public Sub Reset()
    On Error Resume Next
    ' reset control
    Dim mvarFont As StdFont
    Set colHighlights = New Collection
    Set colWorking = New Collection
    Set mvarFont = New StdFont
    mvarFont.Name = "Tahoma"
    mvarFont.Size = 8
    Set Font = mvarFont
    HeaderBold = True
    MonthBold = True
    HeaderFormat = ddd
    BorderStyle = Fixed
    Appearance = ThreeD
    SaturdayIsAWorkingDay = False
    SundayIsAWorkingDay = False
    BackColor = &H8000000F
    CurrentDateBackColor = &HFFC0C0
    CurrentDateBold = True
    DateFormat = "dd/mm/yyyy"
    CurrentDate = Format$(Now, DateFormat)
    NonWorkingBackColor = &HFF80FF
    Width = 3285
    Height = 3000
    HolidayCount = 0
    ReDim Holidays(HolidayCount)
    Err.Clear
End Sub

Private Sub Header_Click(Index As Integer)
    On Error Resume Next
    UserControl.Extender.SetFocus
    Err.Clear
End Sub

Private Sub UserControl_Initialize()
    On Error Resume Next
    'initialize control
    Reset
    Err.Clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Dim mvarFont As StdFont
    Set mvarFont = New StdFont
    mvarFont.Name = "Tahoma"
    mvarFont.Size = 8
    DateFormat = PropBag.ReadProperty("DateFormat", "dd/mm/yyyy")
    Set Font = PropBag.ReadProperty("Font", mvarFont)
    HeaderBold = PropBag.ReadProperty("HeaderBold", True)
    MonthBold = PropBag.ReadProperty("MonthBold", True)
    HeaderFormat = PropBag.ReadProperty("HeaderFormat", ddd)
    BorderStyle = PropBag.ReadProperty("BorderStyle", Fixed)
    Appearance = PropBag.ReadProperty("Appearance", ThreeD)
    BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    CurrentDateBackColor = PropBag.ReadProperty("CurrentDateBackColor", &HFFC0C0)
    CurrentDateBold = PropBag.ReadProperty("CurrentDateBold", False)
    SaturdayIsAWorkingDay = PropBag.ReadProperty("SaturdayIsAWorkingDay", False)
    SundayIsAWorkingDay = PropBag.ReadProperty("SundayIsAWorkingDay", False)
    CurrentDate = PropBag.ReadProperty("CurrentDate", Format$(Now, DateFormat))
    NonWorkingBackColor = PropBag.ReadProperty("NonWorkingBackColor", &HFF80FF)
    UserControl.Refresh
    Err.Clear
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    ' resize the days of the calendar based on calendar size
    Dim eachW As Long
    Dim eachH As Long
    Dim rsCnt As Integer
    
    If UserControl.Height < 3000 Then UserControl.Height = 3000
    If UserControl.Width < 3285 Then UserControl.Width = 3285
    
    ' align horizontally
    eachH = UserControl.Height / 8
    For rsCnt = 1 To 7
        Header(rsCnt).Height = eachH
        Err.Clear
    Next
    
    For rsCnt = 1 To 42
        DayOf(rsCnt).Height = eachH
        Err.Clear
    Next
    
    DayOf(1).Top = Header(1).Height + cmdPreviousDay.Height
    DayOf(8).Top = DayOf(1).Height + DayOf(1).Top
    DayOf(15).Top = DayOf(8).Height + DayOf(8).Top
    DayOf(22).Top = DayOf(15).Height + DayOf(15).Top
    DayOf(29).Top = DayOf(22).Height + DayOf(22).Top
    DayOf(36).Top = DayOf(29).Height + DayOf(29).Top
    
    
    DayOf(2).Top = Header(2).Height + cmdPreviousDay.Height
    DayOf(9).Top = DayOf(2).Height + DayOf(2).Top
    DayOf(16).Top = DayOf(9).Height + DayOf(9).Top
    DayOf(23).Top = DayOf(16).Height + DayOf(16).Top
    DayOf(30).Top = DayOf(23).Height + DayOf(23).Top
    DayOf(37).Top = DayOf(30).Height + DayOf(30).Top
    
    DayOf(3).Top = Header(3).Height + cmdPreviousDay.Height
    DayOf(10).Top = DayOf(3).Height + DayOf(3).Top
    DayOf(17).Top = DayOf(10).Height + DayOf(10).Top
    DayOf(24).Top = DayOf(17).Height + DayOf(17).Top
    DayOf(31).Top = DayOf(24).Height + DayOf(24).Top
    DayOf(38).Top = DayOf(31).Height + DayOf(31).Top
    
    
    DayOf(4).Top = Header(4).Height + cmdPreviousDay.Height
    DayOf(11).Top = DayOf(4).Height + DayOf(4).Top
    DayOf(18).Top = DayOf(11).Height + DayOf(11).Top
    DayOf(25).Top = DayOf(18).Height + DayOf(18).Top
    DayOf(32).Top = DayOf(25).Height + DayOf(25).Top
    DayOf(39).Top = DayOf(32).Height + DayOf(32).Top
    
    DayOf(5).Top = Header(5).Height + cmdPreviousDay.Height
    DayOf(12).Top = DayOf(5).Height + DayOf(5).Top
    DayOf(19).Top = DayOf(12).Height + DayOf(12).Top
    DayOf(26).Top = DayOf(19).Height + DayOf(19).Top
    DayOf(33).Top = DayOf(26).Height + DayOf(26).Top
    DayOf(40).Top = DayOf(33).Height + DayOf(33).Top
    
    DayOf(6).Top = Header(6).Height + cmdPreviousDay.Height
    DayOf(13).Top = DayOf(6).Height + DayOf(6).Top
    DayOf(20).Top = DayOf(13).Height + DayOf(13).Top
    DayOf(27).Top = DayOf(20).Height + DayOf(20).Top
    DayOf(34).Top = DayOf(27).Height + DayOf(27).Top
    DayOf(41).Top = DayOf(34).Height + DayOf(34).Top
    
    DayOf(7).Top = Header(7).Height + cmdPreviousDay.Height
    DayOf(14).Top = DayOf(7).Height + DayOf(7).Top
    DayOf(21).Top = DayOf(14).Height + DayOf(14).Top
    DayOf(28).Top = DayOf(21).Height + DayOf(21).Top
    DayOf(35).Top = DayOf(28).Height + DayOf(28).Top
    DayOf(42).Top = DayOf(35).Height + DayOf(35).Top
    
    ' align vertically
    eachW = UserControl.Width / 7
    For rsCnt = 1 To 7
        Header(rsCnt).Width = eachW - 10
        Err.Clear
    Next
    For rsCnt = 2 To 7
        Header(rsCnt).Left = Header(rsCnt - 1).Width + Header(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 1 To 42
        DayOf(rsCnt).Width = eachW - 10
        Err.Clear
    Next
    
    For rsCnt = 2 To 7
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 9 To 14
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 16 To 21
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 23 To 28
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 30 To 35
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    For rsCnt = 37 To 42
        DayOf(rsCnt).Left = DayOf(rsCnt - 1).Width + DayOf(rsCnt - 1).Left
        Err.Clear
    Next
    
    txtMonth.Left = cmdPreviousDay.Width + cmdPreviousYear.Width
    txtMonth.Width = UserControl.Width - (cmdPreviousDay.Width + cmdPreviousYear.Width + cmdNextYear.Width + cmdNextDay.Width) - 40
    cmdNextYear.Left = txtMonth.Width + txtMonth.Left
    cmdNextDay.Left = cmdNextYear.Left + cmdNextYear.Width
    
    'UserControl.Height = cmdPreviousDay.Height + Header(1).Height + DayOf(1).Height + DayOf(8).Height + DayOf(15).Height + DayOf(22).Height + DayOf(29).Height + DayOf(36).Height
    Err.Clear
End Sub

Public Property Get WorkingDays() As Integer
    On Error Resume Next
    Dim rsCnt As Integer
    Dim workCnt As Integer
    
    workCnt = 0
    For rsCnt = 1 To 42
        Select Case DayOf(rsCnt).BackColor
        Case &H80000005, mvarCurrentDateBackColor
            workCnt = workCnt + 1
        End Select
    Next
    WorkingDays = workCnt
End Property

Public Property Get Appearance() As AppearanceEnum
    On Error Resume Next
    Appearance = mvarAppearance
    Err.Clear
End Property

Public Property Let Appearance(ByVal Value As AppearanceEnum)
    On Error Resume Next
    ' set appearance of the calendar
    UserControl.Appearance = Value
    mvarAppearance = Value
    PropertyChanged "Appearance"
    UserControl.Refresh
    Err.Clear
End Property

Public Property Get NonWorkingBackColor() As OLE_COLOR
    On Error Resume Next
    NonWorkingBackColor = mvarNonWorkingBackColor
    Err.Clear
End Property

Public Property Let NonWorkingBackColor(ByVal Value As OLE_COLOR)
    On Error Resume Next
    ' set the color for the non working days
    Dim rsCnt As Integer
    mvarNonWorkingBackColor = Value
    For rsCnt = 1 To 42
        Select Case rsCnt
        Case 7, 14, 21, 28, 35, 42
            ' saturday
            DayOf(rsCnt).BackColor = IIf(mvarSaturdayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        Case 1, 8, 15, 22, 29, 36
            ' sunday
            DayOf(rsCnt).BackColor = IIf(mvarSundayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        End Select
        If Len(DayOf(rsCnt)) = 0 Then
            DayOf(rsCnt).BackColor = UserControl.BackColor
            DayOf(rsCnt).Enabled = False
        End If
        Err.Clear
    Next
    
    DayOf(mvarCurrentDay + sPos - 1).BackColor = mvarCurrentDateBackColor
    DayOf(mvarCurrentDay + sPos - 1).Font.Bold = mvarCurrentDateBold
    
    PropertyChanged "NonWorkingBackColor"
    Err.Clear
End Property

Public Property Get CurrentDateBackColor() As OLE_COLOR
    On Error Resume Next
    CurrentDateBackColor = mvarCurrentDateBackColor
    Err.Clear
End Property

Public Property Get StartDate() As String
    On Error Resume Next
    StartDate = Format$(mvarStartDate, DateFormat)
    Err.Clear
End Property

Public Property Get EndDate() As String
    On Error Resume Next
    EndDate = Format$(mvarEndDate, DateFormat)
    Err.Clear
End Property


Public Property Let CurrentDateBackColor(ByVal Value As OLE_COLOR)
    On Error Resume Next
    ' set the color of the current date
    mvarCurrentDateBackColor = Value
    DayOf(mvarCurrentDay + sPos - 1).BackColor = mvarCurrentDateBackColor
    PropertyChanged "CurrentDateBackColor"
    Err.Clear
End Property


Public Property Get BackColor() As OLE_COLOR
    On Error Resume Next
    BackColor = mvarBackColor
    Err.Clear
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
    On Error Resume Next
    ' set the back color of the calendar
    Dim rsCnt As Integer
    mvarBackColor = Value
    UserControl.BackColor = Value
    For rsCnt = 1 To 42
        If Len(DayOf(rsCnt)) = 0 Then
            DayOf(rsCnt).BackColor = UserControl.BackColor
            DayOf(rsCnt).Enabled = False
        End If
        Err.Clear
    Next
    PropertyChanged "BackColor"
    Err.Clear
End Property

Public Property Get Font() As Font
    On Error Resume Next
    Set Font = UserControl.Font
    Err.Clear
End Property

Public Property Set Font(ByVal New_Font As Font)
    On Error Resume Next
    ' set the font of the calendar
    Dim rsCnt As Integer
    Set UserControl.Font = New_Font
    For rsCnt = 1 To 7
        Header(rsCnt).Font.Name = New_Font.Name
        Err.Clear
    Next
    For rsCnt = 1 To 42
        DayOf(rsCnt).Font.Name = New_Font.Name
        Err.Clear
    Next
    cmdNextDay.Font.Name = New_Font.Name
    cmdNextYear.Font.Name = New_Font.Name
    cmdPreviousYear.Font.Name = New_Font.Name
    cmdPreviousDay.Font.Name = New_Font.Name
    txtMonth.Font.Name = New_Font.Name
    PropertyChanged "Font"
    Err.Clear
End Property

Public Property Get CurrentDate() As String
    On Error Resume Next
    CurrentDate = Format$(mvarCurrentDate, DateFormat)
    Err.Clear
End Property

Public Property Let CurrentDate(ByVal New_Date As String)
    On Error Resume Next
    ' set the current date of the calendar
    Dim sDate As String
    Dim eDate As String
    Dim sDates As String
    Dim numDays As Integer
    Dim rsCnt As Integer
    'Dim sDay As String
    Dim xDay As Integer
    
    mvarCurrentDate = New_Date
    txtMonth.Text = Format$(New_Date, "mmmm yyyy")
    mvarCurrentYear = Year(New_Date)
    mvarCurrentMonth = Month(New_Date)
    mvarCurrentDay = Day(New_Date)
    ClearDays
    
    ' find the starting date and ending date for the month
    sDates = StartEndDate(Format$(New_Date, "yyyymm"), "b")
    sDate = Split(sDates, ",")(0)
    eDate = Split(sDates, ",")(1)
    mvarStartDate = sDate
    mvarEndDate = eDate
    
    ' calculate the number of days in the month
    numDays = DateDiff("d", sDate, eDate)
    ' determine the starting position
    Select Case Format$(sDate, "ddd")
    Case "Sun"
        sPos = 1
    Case "Mon"
        sPos = 2
    Case "Tue"
        sPos = 3
    Case "Wed"
        sPos = 4
    Case "Thu"
        sPos = 5
    Case "Fri"
        sPos = 6
    Case "Sat"
        sPos = 7
    End Select
    
    ' put day on days
    xDay = 1
    For rsCnt = sPos To numDays + sPos
        DayOf(rsCnt).Text = xDay
        DayOf(rsCnt).Tag = Format$(xDay & "/" & mvarCurrentMonth & "/" & mvarCurrentYear, DateFormat)
        DayOf(rsCnt).ToolTipText = ""
        xDay = xDay + 1
        Err.Clear
    Next
    
    ' change background of each empty text
    For rsCnt = 1 To 42
        Select Case rsCnt
        Case 7, 14, 21, 28, 35, 42
            ' saturday
            DayOf(rsCnt).BackColor = IIf(mvarSaturdayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        Case 1, 8, 15, 22, 29, 36
            ' sunday
            DayOf(rsCnt).BackColor = IIf(mvarSundayIsAWorkingDay = True, &H80000005, mvarNonWorkingBackColor)
        End Select
        If Len(DayOf(rsCnt)) = 0 Then
            DayOf(rsCnt).BackColor = UserControl.BackColor
            DayOf(rsCnt).Enabled = False
        End If
        Err.Clear
    Next
    
    DayOf(mvarCurrentDay + sPos - 1).BackColor = mvarCurrentDateBackColor
    DayOf(mvarCurrentDay + sPos - 1).Font.Bold = mvarCurrentDateBold
    
    MarkHolidays
    MarkHighlights
    MarkWorkingDays
    PropertyChanged "CurrentDate"
    UserControl.Refresh
    Err.Clear
End Property

Private Function StartEndDate(ByVal Yyyymm As String, Optional ByVal Sread As String = "") As String
    On Error Resume Next
    ' find the start and end date for a month
    Dim smm As String
    Dim syy As String
    Dim sDate As String
    Dim eyy As String
    Dim emm As String
    Dim eDate As String
    Dim dLen As Integer
    If Len(Sread) = 0 Then
        Sread = "B"
    End If
    dLen = Len(Yyyymm) - 2
    smm = Right$(Yyyymm, 2)
    syy = Left$(Yyyymm, dLen)
    sDate = "01/" & smm & "/" & syy
    emm = Val(smm) + 1
    eyy = syy
    Select Case Val(emm)
    Case Is > 12
        emm = Val(emm) - 12
        eyy = Val(eyy) + 1
        If Val(eyy) = 100 Then
            eyy = "00"
        End If
    End Select
    eDate = "01/" & emm & "/" & eyy
    eDate = CDate(DateAdd("d", -1, CDate(eDate)))
    sDate = Format$(sDate, "dd/mm/yyyy")
    eDate = Format$(eDate, "dd/mm/yyyy")
    Select Case UCase$(Left$(Sread, 1))
    Case "S"
        StartEndDate = sDate
    Case "E"
        StartEndDate = eDate
    Case "B"
        StartEndDate = sDate & "," & eDate
    End Select
    Err.Clear
End Function

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    Dim mvarFont As StdFont
    Set mvarFont = New StdFont
    mvarFont.Name = "Tahoma"
    mvarFont.Size = 8
    PropBag.WriteProperty "DateFormat", mvarDateFormat, "dd/mm/yyyy"
    PropBag.WriteProperty "Font", mvarFont
    PropBag.WriteProperty "HeaderBold", mvarHeaderBold, True
    PropBag.WriteProperty "MonthBold", mvarMonthBold, True
    PropBag.WriteProperty "HeaderFormat", mvarHeaderFormat, ddd
    PropBag.WriteProperty "BorderStyle", mvarBorderStyle, Fixed
    PropBag.WriteProperty "Appearance", mvarAppearance, ThreeD
    PropBag.WriteProperty "CurrentDate", mvarCurrentDate, Format$(Now, DateFormat)
    PropBag.WriteProperty "BackColor", mvarBackColor, &H8000000F
    PropBag.WriteProperty "CurrentDateBackColor", mvarCurrentDateBackColor, &HFFC0C0
    PropBag.WriteProperty "CurrentDateBold", mvarCurrentDateBold, False
    PropBag.WriteProperty "SaturdayIsAWorkingDay", mvarSaturdayIsAWorkingDay, False
    PropBag.WriteProperty "SundayIsAWorkingDay", mvarSundayIsAWorkingDay, False
    PropBag.WriteProperty "NonWorkingBackColor", mvarNonWorkingBackColor, &HFF80FF
    Err.Clear
End Sub
