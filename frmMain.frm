VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "2D Geometry Tools"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFILLET 
      Caption         =   "Fillet"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      ToolTipText     =   "Find Arc (of a given radius) tangent to two lines"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdFowlerAngle 
      Caption         =   "Fast Atan2 approx"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCircleLine 
      Caption         =   "Circle Line Intersection"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdAngleDiff 
      Caption         =   "Angle Difference"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdNearestOnLine 
      Caption         =   "Nearest point on Line"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdMkCircle 
      Caption         =   "Circle from 3 Points"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton cmdTangentTwoCircles 
      Caption         =   "Tangents of Two Circles"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdCircleCircle 
      Caption         =   "Circle Circle Intersection"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton cmdVectorReflect 
      Caption         =   "Vector Reflect"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdVectorProject 
      Caption         =   "Vector Project"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdLineLine 
      Caption         =   "Line Line Intersection"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New  --->"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************
' AUTHOR: Roberto Mior
' reexre@gmail.com
' Suggestions or new Tools are wellcome!
' Most Function taken from http://paulbourke.net/geometry
'***********************************************************************************


Private Sub cmdAngleDiff_Click()
    Load frmAngleDiff
    frmAngleDiff.Show

End Sub

Private Sub cmdCircleCircle_Click()
    Load frmCircleCircle
    frmCircleCircle.Show

End Sub

Private Sub cmdCircleLine_Click()
    Load frmCircleLine
    frmCircleLine.Show
End Sub

Private Sub cmdFILLET_Click()
    Load frmFILLET
    frmFILLET.Show

End Sub

Private Sub cmdFowlerAngle_Click()
    Load frmFASTAtan2
    frmFASTAtan2.Show
End Sub

Private Sub cmdLineLine_Click()
    Load frmLineLine
    frmLineLine.Show

End Sub

Private Sub cmdMkCircle_Click()
    Load frmMKcircleBy3Points
    frmMKcircleBy3Points.Show

End Sub

Private Sub cmdNearestOnLine_Click()
    Load frmNearestFromLine
    frmNearestFromLine.Show

End Sub

Private Sub cmdTangentTwoCircles_Click()
    Load frmTangentTwoCircles
    frmTangentTwoCircles.Show

End Sub

Private Sub cmdVectorProject_Click()
    Load frmVectorProject
    frmVectorProject.Show
End Sub

Private Sub cmdVectorReflect_Click()
    Load frmVectorReflect
    frmVectorReflect.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
