VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "RTrimZZ (safe array)"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   2895
   End
   Begin VB.ComboBox cmbTest 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   240
      Width           =   8055
   End
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "LTrimZ (safe array)"
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   2895
   End
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "RTrimZ (safe array)"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "TrimZZ_2 (InStrB)"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "TrimZZ (InStr)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton cmdTrimNull 
      Caption         =   "TrimZ (lstrlenW)"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   12
      Top             =   3240
      Width           =   5055
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   9
      Top             =   2760
      Width           =   5055
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   1800
      Width           =   5055
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label lblTrimNull 
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ITERATIONS = 100000
Dim TEST As String

Private Sub cmbTest_Click()
    Dim I As Long
    TEST = Replace$(cmbTest.List(cmbTest.ListIndex), "Z", vbNullChar)
    For I = 0 To lblTrimNull.UBound
        lblTrimNull(I).Caption = "<- Untested"
    Next I
End Sub

Private Sub cmdTrimNull_Click(Index As Integer)
    Dim C As Long, R As String, T As Double
    C = ITERATIONS
    Timing = 0
    Select Case Index
        Case 0: For C = C To 1 Step -1: R = TrimZ(TEST): Next
        Case 1: For C = C To 1 Step -1: R = TrimZZ(TEST): Next
        Case 2: For C = C To 1 Step -1: R = TrimZZ_2(TEST): Next
        Case 3: For C = C To 1 Step -1: R = RTrimZ(TEST): Next
        Case 4: For C = C To 1 Step -1: R = LTrimZ(TEST): Next
        Case 5: For C = C To 1 Step -1: R = RTrimZZ(TEST): Next
    End Select
    T = Timing
    lblTrimNull(Index).Caption = ITERATIONS & " loops @ " & Format$(T * 1000, "0.00000") & " ms, result: """ & Replace$(R, vbNullChar, "[Z]") & """"
End Sub

Private Sub Form_Load()
    cmbTest.AddItem "TestingZZZZZ"
    cmbTest.AddItem "ZZZZZTesting"
    cmbTest.AddItem "ZZZZZTestingZZZZZ"
    cmbTest.AddItem "TestingZTestingZ"
    cmbTest.AddItem "TestingZTestingZZ"
    cmbTest.AddItem "ZTestingZTesting"
    cmbTest.AddItem "ZZTestingZTesting"
    cmbTest.AddItem String$(1022, "A") & "ZZ"
    cmbTest.AddItem "AA" & String$(1022, "ZZ")
    cmbTest.ListIndex = 0
End Sub
