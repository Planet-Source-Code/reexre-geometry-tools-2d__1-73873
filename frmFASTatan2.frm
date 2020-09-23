VERSION 5.00
Begin VB.Form frmFASTAtan2 
   Caption         =   "Fast Atan2 Approximation"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSPEED 
      Caption         =   "Test Speed"
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmFASTatan2.frx":0000
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox TXT 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmFASTatan2.frx":0143
      Top             =   1560
      Width           =   6135
   End
   Begin VB.Line FAline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   344
      X2              =   408
      Y1              =   344
      Y2              =   392
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   344
      X2              =   360
      Y1              =   344
      Y2              =   280
   End
End
Attribute VB_Name = "frmFASTAtan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Pmove          As Long

Private Sub cmdSPEED_Click()
    Dim R          As Double

    Dim I          As Long
    Dim T
    Dim T1
    Dim T2
    Dim A          As Single

    Me.WindowState = vbMinimized
    DoEvents
    MsgBox "Will be performed a test speed.... Wait....."

    R = Rnd

    Randomize R
    T = Timer
    For I = 0 To 8000000
        A = Atan2(Rnd * 20 - 10, Rnd * 20 - 10)
    Next
    T1 = Timer - T

    Randomize R
    T = Timer
    For I = 0 To 8000000
        A = Atan2Fast2(Rnd * 20 - 10, Rnd * 20 - 10)
    Next
    T2 = Timer - T

    MsgBox "Time For Atan2 = " & T1 & vbCrLf & "Time For Atan2Fast2 = " & T2 & vbCrLf & vbCrLf & _
           IIf(T1 > T2, "Fast Atan2 is " & T1 / T2 & " Times Faster", "Atan2 is " & T2 / T1 & " Times Faster")

    Me.WindowState = vbNormal

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dx         As Single
    Dim dy         As Single
    Dim AT2        As Single
    Dim FA         As Single
    Dim AT2_360    As Single
    Dim FA_360     As Single
    Dim D          As Single

    If Pmove Then
        Line1.x2 = X
        Line1.y2 = Y

        dx = Line1.x2 - Line1.x1
        dy = Line1.y2 - Line1.y1
        D = Sqr(dx * dx + dy * dy)

        AT2 = Atan2(dx, dy)
        FA = Atan2Fast2(dx, dy)

        If AT2 < 0 Then AT2 = AT2 + PI2
        AT2_360 = Format(360 * AT2 / PI2, "000.0")
        FA_360 = Format(360 * FA / PI2, "000.0")

        TXT = "dx = " & dx & vbCrLf
        TXT = TXT & "dy = " & dy & vbCrLf & vbCrLf

        TXT = TXT & "      Atan2 = " & AT2 & "  " & AT2_360 & "°" & vbCrLf & _
              "Atan2Fast2 = " & FA & "  " & FA_360 & "°   [" & Format(4 * FA / PI, "0.000") & "]"


        FAline.x2 = FAline.x1 + Cos(FA) * D
        FAline.y2 = FAline.y1 + Sin(FA) * D


    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0

End Sub

