VERSION 5.00
Begin VB.Form frmCircleLine 
   Caption         =   "Circle Line Intersection"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Res2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   255
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Move Circle or  Line  with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Shape Res1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   255
      Left            =   7320
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   255
   End
   Begin VB.Shape Circle1 
      Height          =   1935
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Line Line1 
      X1              =   200
      X2              =   464
      Y1              =   208
      Y2              =   176
   End
End
Attribute VB_Name = "frmCircleLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ci1        As geoCircle

'Private C As geoCircle
Private Inter      As geoPointVector2D
Private Pmove      As Long

Private L1         As geoLine


Private Sub Form_Load()
    Circle1.Height = Circle1.Width

    L1.P1.X = Line1.x1
    L1.P1.Y = Line1.y1
    L1.P2.X = Line1.x2
    L1.P2.Y = Line1.y2

    FindCircleLine
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 3)  As Single
    Dim Dmin       As Single
    Dim I          As Long
    D(1) = DistFromPoint2(Ci1.Center, X, Y)
    D(2) = DistFromPoint2(L1.P1, X, Y)
    D(3) = DistFromPoint2(L1.P2, X, Y)

    Dmin = 999999999999#
    For I = 1 To UBound(D)
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pmove <> 0 Then
        Select Case Pmove
            Case 1
                Circle1.Left = X - Circle1.Width \ 2
                Circle1.Top = Y - Circle1.Height \ 2
            Case 2
                Line1.x1 = X
                Line1.y1 = Y
            Case 3
                Line1.x2 = X
                Line1.y2 = Y

        End Select

        FindCircleLine
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindCircleLine()
    Dim S1         As geoPointVector2D
    Dim S2         As geoPointVector2D


    Ci1.Center.X = Circle1.Left + Circle1.Width \ 2
    Ci1.Center.Y = Circle1.Top + Circle1.Height \ 2
    Ci1.Radius = Circle1.Width \ 2
    L1.P1.X = Line1.x1
    L1.P1.Y = Line1.y1
    L1.P2.X = Line1.x2
    L1.P2.Y = Line1.y2

    IntersectCircleLine Ci1, L1, S1, S2

    Res1.Left = S1.X - Res1.Width \ 2
    Res1.Top = S1.Y - Res1.Width \ 2
    Res2.Left = S2.X - Res2.Width \ 2
    Res2.Top = S2.Y - Res2.Width \ 2

    If S1.Bool = True Then Res1.BorderColor = vbBlue Else: Res1.BorderColor = vbRed
    If S2.Bool = True Then Res2.BorderColor = vbBlue Else: Res2.BorderColor = vbRed

End Sub



