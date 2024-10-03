VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LAB ATTENDANCE MANEGEMENT"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "All Attendance details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   8
      Top             =   5640
      Width           =   3975
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2520
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4445
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   121503745
      CurrentDate     =   45465
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Attendance details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   4680
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Present"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   4
      Top             =   5520
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   2160
      Left            =   5400
      TabIndex        =   3
      Top             =   1320
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      Height          =   5775
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Enter Specific Roll No."
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Roll No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lab Attendance "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim attendance(1 To 100) As Collection

Private Sub Command3_Click()
Dim i As Integer
Dim hasAttendance As Boolean
List1.Clear

For i = 1 To 100
If attendance(i).Count > 0 Then
        ' Loop through each recorded attendance date for roll number i
For Each recordedDate In attendance(i)
            ' Format the date as desired (adjust format as needed)
Dim formattedDate As String
formattedDate = Format(recordedDate, "dd/mm/yyyy")
Dim message As String
message = "Roll Number " & i & " has recorded attendance on " & formattedDate
        
List1.AddItem Trim(message)
hasAttendance = True
Next recordedDate
End If
Next i

If Not hasAttendance Then
List1.AddItem "No attendance recorded for any roll number."
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Dim i As Integer
For i = 1 To 100
Set attendance(i) = New Collection
Next i
End Sub
Private Sub Command1_Click()
Dim rollNumber As Integer
Dim currentDate As Date
Dim dateItem As Variant
rollNumber = Val(Text1.Text)
currentDate = MonthView1.Value
If rollNumber < 1 Or rollNumber > 100 Then
MsgBox "Invalid roll number. Please enter a number between 1 and 100."
Exit Sub
End If
For Each dateItem In attendance(rollNumber)
If dateItem = currentDate Then
MsgBox "Attendance already marked for this date"
Exit Sub
End If
Next dateItem
attendance(rollNumber).Add currentDate
MsgBox "Attendance marked for roll number: " & rollNumber & " on " & currentDate
End Sub
Private Sub Command2_Click()
Dim rollNumber As Integer
Dim presentDays As Integer
Dim dateItem As Variant
rollNumber = Val(Text2.Text)
If rollNumber < 1 Or rollNumber > 100 Then
MsgBox "Invalid roll number. Please enter a number between 1 and 100."
Exit Sub
End If
List1.Clear
presentDays = attendance(rollNumber).Count
List1.AddItem "Roll Number: " & rollNumber & " is present for " & presentDays & " days."
For Each dateItem In attendance(rollNumber)
List1.AddItem dateItem
Next dateItem
End Sub

