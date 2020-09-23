VERSION 5.00
Begin VB.Form frmTangentTwoCircles 
   Caption         =   "Tangent Two Circles"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   224
      X2              =   472
      Y1              =   280
      Y2              =   344
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   200
      X2              =   392
      Y1              =   312
      Y2              =   376
   End
   Begin VB.Shape C2 
      BorderWidth     =   2
      Height          =   1935
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Shape C1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   2535
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Move Circles with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTangentTwoCircles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ci1        As geoCircle
Private Ci2        As geoCircle
'Private C As geoCircle
Private Inter      As geoPointVector2D
Private Pmove      As Long

Private L1         As geoLine
Private L2         As geoLine


Private Sub Form_Load()
    C1.Height = C1.Width
    C2.Height = C2.Width

    FindTangets
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

        FindTangets
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindTangets()


    Ci1.Center.X = C1.Left + C1.Width \ 2
    Ci1.Center.Y = C1.Top + C1.Height \ 2
    Ci2.Center.X = C2.Left + C2.Width \ 2
    Ci2.Center.Y = C2.Top + C2.Height \ 2

    Ci1.Radius = C1.Width \ 2
    Ci2.Radius = C2.Width \ 2


    TangentTwoCircles Ci1, Ci2, L1, L2

    'Stop

    If L1.P1.Bool Or L1.P2.Bool Then
        Line1.x1 = L1.P1.X
        Line1.y1 = L1.P1.Y
        Line1.x2 = L1.P2.X
        Line1.y2 = L1.P2.Y
        Line1.Visible = True
    Else
        Line1.Visible = False
    End If

    If L2.P1.Bool Or L2.P2.Bool Then
        Line2.x1 = L2.P1.X
        Line2.y1 = L2.P1.Y
        Line2.x2 = L2.P2.X
        Line2.y2 = L2.P2.Y
        Line2.Visible = True
    Else
        Line2.Visible = False
    End If

End Sub


