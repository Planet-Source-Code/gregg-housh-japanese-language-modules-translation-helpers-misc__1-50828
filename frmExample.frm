VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExample 
   Caption         =   "Japanese Modules Example"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Numbers"
      TabPicture(0)   =   "frmExample.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNumbers"
      Tab(0).Control(1)=   "lblNumbers2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Verbs"
      TabPicture(1)   =   "frmExample.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblVerb"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblVerb2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblVerb3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Dates && Times"
      TabPicture(2)   =   "frmExample.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDates"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblDates2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblDates3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Adjectives"
      TabPicture(3)   =   "frmExample.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblAdjectives"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblAdjectives2"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblAdjectives3"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.Label lblAdjectives3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -69600
         TabIndex        =   11
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAdjectives2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -72240
         TabIndex        =   10
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAdjectives 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblDates3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -69600
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblDates2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -72240
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblDates 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblVerb3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   5280
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblVerb2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblVerb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblNumbers2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -72360
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label lblNumbers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Label Label 
      Caption         =   "These examples are just some simple uses of the modules.  They show what the different functions can do."
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   7695
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' have to call this to fill the gVerbExamples array.
    SetupVerbExamples
    
    Dim strNumbers As String
    Dim strVerbs As String
    Dim iLoop As Integer
    Dim strDates As String
    Dim strAdjectives As String
    
    strNumbers = "0 - 9" & vbCrLf & vbCrLf
    strNumbers = strNumbers & "0 = " & jNumber(0) & vbCrLf
    strNumbers = strNumbers & "1 = " & jNumber(1) & vbCrLf
    strNumbers = strNumbers & "2 = " & jNumber(2) & vbCrLf
    strNumbers = strNumbers & "3 = " & jNumber(3) & vbCrLf
    strNumbers = strNumbers & "4 = " & jNumber(4) & vbCrLf
    strNumbers = strNumbers & "5 = " & jNumber(5) & vbCrLf
    strNumbers = strNumbers & "6 = " & jNumber(6) & vbCrLf
    strNumbers = strNumbers & "7 = " & jNumber(7) & vbCrLf
    strNumbers = strNumbers & "8 = " & jNumber(8) & vbCrLf
    strNumbers = strNumbers & "9 = " & jNumber(9)
    lblNumbers.Caption = strNumbers
    
    strNumbers = "misc" & vbCrLf & vbCrLf
    strNumbers = strNumbers & "10 = " & jNumber(10) & vbCrLf
    strNumbers = strNumbers & "100 = " & jNumber(100) & vbCrLf
    strNumbers = strNumbers & "1000 = " & jNumber(1000) & vbCrLf
    strNumbers = strNumbers & "10000 = " & jNumber(10000) & vbCrLf
    strNumbers = strNumbers & "100000 = " & jNumber(100000) & vbCrLf
    strNumbers = strNumbers & "1000000 = " & jNumber(1000000) & vbCrLf
    strNumbers = strNumbers & "10000000 = " & jNumber(10000000) & vbCrLf
    strNumbers = strNumbers & "100000000 = " & jNumber(100000000) & vbCrLf
    strNumbers = strNumbers & vbCrLf & "More Misc" & vbCrLf & vbCrLf
    strNumbers = strNumbers & "11 = " & jNumber(11) & vbCrLf
    strNumbers = strNumbers & "213 = " & jNumber(213) & vbCrLf
    strNumbers = strNumbers & "4387 = " & jNumber(4387) & vbCrLf
    strNumbers = strNumbers & "97345 = " & jNumber(97345) & vbCrLf
    strNumbers = strNumbers & "583240 = " & jNumber(583240) & vbCrLf
    lblNumbers2.Caption = strNumbers
    
    strVerbs = "taberu = to eat" & vbCrLf
    For iLoop = 3 To 15
        strVerbs = strVerbs & ConjugateVerb("taberu", iLoop) & vbCrLf
        lblVerb.Caption = strVerbs
    Next iLoop
    strVerbs = ""
    For iLoop = 16 To 29
        strVerbs = strVerbs & ConjugateVerb("taberu", iLoop) & vbCrLf
        lblVerb2.Caption = strVerbs
    Next iLoop
    strVerbs = ""
    For iLoop = 30 To 37
        strVerbs = strVerbs & ConjugateVerb("taberu", iLoop) & vbCrLf
        lblVerb3.Caption = strVerbs
    Next iLoop
    
    strDates = "Months" & vbCrLf & vbCrLf
    strDates = strDates & "January = " & jMonth(CDate("1/1/04")) & vbCrLf
    strDates = strDates & "Febuary = " & jMonth(CDate("2/1/04")) & vbCrLf
    strDates = strDates & "March = " & jMonth(CDate("3/1/04")) & vbCrLf
    strDates = strDates & "April = " & jMonth(CDate("4/1/04")) & vbCrLf
    strDates = strDates & "May = " & jMonth(CDate("5/1/04")) & vbCrLf
    strDates = strDates & "June = " & jMonth(CDate("6/1/04")) & vbCrLf
    strDates = strDates & "July = " & jMonth(CDate("7/1/04")) & vbCrLf
    strDates = strDates & "August = " & jMonth(CDate("8/1/04")) & vbCrLf
    strDates = strDates & "September = " & jMonth(CDate("9/1/04")) & vbCrLf
    strDates = strDates & "October = " & jMonth(CDate("10/1/04")) & vbCrLf
    strDates = strDates & "November = " & jMonth(CDate("11/1/04")) & vbCrLf
    strDates = strDates & "December = " & jMonth(CDate("12/1/04")) & vbCrLf
    lblDates.Caption = strDates
    
    strDates = "Days of the week" & vbCrLf & vbCrLf
    strDates = strDates & "Sunday = " & jWeekDay(CDate("1/4/04")) & vbCrLf
    strDates = strDates & "Monday = " & jWeekDay(CDate("1/5/04")) & vbCrLf
    strDates = strDates & "Tuesday = " & jWeekDay(CDate("1/6/04")) & vbCrLf
    strDates = strDates & "Wednesday = " & jWeekDay(CDate("1/7/04")) & vbCrLf
    strDates = strDates & "Thursday = " & jWeekDay(CDate("1/8/04")) & vbCrLf
    strDates = strDates & "Friday = " & jWeekDay(CDate("1/9/04")) & vbCrLf
    strDates = strDates & "Saturday = " & jWeekDay(CDate("1/10/04")) & vbCrLf
    lblDates2.Caption = strDates

    strDates = "Times" & vbCrLf & vbCrLf
    strDates = strDates & "1 AM = " & jHour(CDate("1:00 AM")) & vbCrLf
    strDates = strDates & "2 AM = " & jHour(CDate("2:00 AM")) & vbCrLf
    strDates = strDates & "3 AM = " & jHour(CDate("3:00 AM")) & vbCrLf
    strDates = strDates & "4 AM = " & jHour(CDate("4:00 AM")) & vbCrLf
    strDates = strDates & "5 AM = " & jHour(CDate("5:00 AM")) & vbCrLf
    strDates = strDates & "6 AM = " & jHour(CDate("6:00 AM")) & vbCrLf
    strDates = strDates & "7 AM = " & jHour(CDate("7:00 AM")) & vbCrLf
    strDates = strDates & "8 AM = " & jHour(CDate("8:00 AM")) & vbCrLf
    strDates = strDates & "9 AM = " & jHour(CDate("9:00 AM")) & vbCrLf
    strDates = strDates & "10 AM = " & jHour(CDate("10:00 AM")) & vbCrLf
    strDates = strDates & "11 AM = " & jHour(CDate("11:00 AM")) & vbCrLf
    strDates = strDates & "12 AM = " & jHour(CDate("12:00 AM")) & vbCrLf & vbCrLf
    strDates = strDates & "For PM change gozen to gogo"
    lblDates3.Caption = strDates
    
    strAdjectives = "omoshiroi = interesting" & vbCrLf & vbCrLf
    For iLoop = 1 To 10
        strAdjectives = strAdjectives & ConjugateAdjective("omoshiroi", iLoop) & vbCrLf
    Next iLoop
    lblAdjectives.Caption = strAdjectives
    strAdjectives = ""
    For iLoop = 11 To 25
        strAdjectives = strAdjectives & ConjugateAdjective("omoshiroi", iLoop) & vbCrLf
    Next iLoop
    lblAdjectives2.Caption = strAdjectives
    strAdjectives = ""
    For iLoop = 26 To 32
        strAdjectives = strAdjectives & ConjugateAdjective("omoshiroi", iLoop) & vbCrLf
    Next iLoop
    lblAdjectives3.Caption = strAdjectives
    
End Sub
