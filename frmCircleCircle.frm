VERSION 5.00
Begin VB.Form frmCircleCircle 
   Caption         =   "Circle Circle Intersection"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Move Circles with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Sol2 
      FillColor       =   &H0000C000&
      Height          =   255
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Shape Sol1 
      FillColor       =   &H0000C000&
      Height          =   255
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   255
   End
   Begin VB.Shape C2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2535
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Shape C1 
      BorderWidth     =   2
      Height          =   1935
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmCircleCircle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ci1        As geoCircle
Private Ci2        As geoCircle
'Private C As geoCircle
Private Inter      As geoPointVector2D
Private Pmove      As Long

Private P1         As geoPointVector2D
Private P2         As geoPointVector2D


Private Sub Form_Load()
    C1.Height = C1.Width
    C2.Height = C2.Width

    FindCircCirc
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 2)  As Single
    Dim Dmin       As Single
    Dim I          As Long
    D(1) = DistFromPoint2(Ci1.Center, X, Y)
    D(2) = DistFromPoint2(Ci2.Center, X, Y)

    Dmin = 999999999999#
    For I = 1 To 2
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pmove <> 0 Then
        Select Case Pmove
            Case 1
                C1.Left = X - C1.Width \ 2
                C1.Top = Y - C1.Height \ 2
            Case 2
                C2.Left = X - C2.Width \ 2
                C2.Top = Y - C2.Height \ 2
        End Select

        FindCircCirc
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindCircCirc()


    Ci1.Center.X = C1.Left + C1.Width \ 2
    Ci1.Center.Y = C1.Top + C1.Height \ 2
    Ci2.Center.X = C2.Left + C2.Width \ 2
    Ci2.Center.Y = C2.Top + C2.Height \ 2

    Ci1.Radius = C1.Width \ 2
    Ci2.Radius = C2.Width \ 2


    IntersectOfCircles Ci1, Ci2, P1, P2

    If P1.Bool Then
        Sol1.Left = P1.X - Sol1.Width \ 2
        Sol1.Top = P1.Y - Sol1.Height \ 2
        Sol1.Visible = True
    Else
        Sol1.Visible = False
    End If

    If P2.Bool Then
        Sol2.Left = P2.X - Sol2.Width \ 2
        Sol2.Top = P2.Y - Sol2.Height \ 2
        Sol2.Visible = True
    Else
        Sol2.Visible = False
    End If


End Sub

