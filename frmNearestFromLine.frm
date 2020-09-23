VERSION 5.00
Begin VB.Form frmNearestFromLine 
   Caption         =   "Nearest Point on Line"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Move Line or point with Mouse"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape SolPoint 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   255
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Point1 
      Height          =   255
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   192
      X2              =   416
      Y1              =   216
      Y2              =   88
   End
End
Attribute VB_Name = "frmNearestFromLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private L1         As geoLine
Private Inter      As geoPointVector2D
Private Pmove      As Long

Private P1         As geoPointVector2D
Private SolP       As geoPointVector2D


Private Sub Form_Load()


    FindNearest
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 3)  As Single
    Dim Dmin       As Single
    Dim I          As Long
    D(1) = DistFromPoint2(mkPoint(Line1.x1, Line1.y1), X, Y)
    D(2) = DistFromPoint2(mkPoint(Line1.x2, Line1.y2), X, Y)
    D(3) = DistFromPoint2(mkPoint(Point1.Left + Point1.Width \ 2, Point1.Top + Point1.Width \ 2), X, Y)

    Dmin = 999999999999#
    For I = 1 To 3
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pmove <> 0 Then
        Select Case Pmove
            Case 1
                Line1.x1 = X
                Line1.y1 = Y
            Case 2
                Line1.x2 = X
                Line1.y2 = Y
            Case 3
                Point1.Left = X - Point1.Width \ 2
                Point1.Top = Y - Point1.Width \ 2
        End Select

        FindNearest
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindNearest()

    L1.P1.X = Line1.x1
    L1.P1.Y = Line1.y1
    L1.P2.X = Line1.x2
    L1.P2.Y = Line1.y2

    P1.X = Point1.Left + Point1.Width \ 2
    P1.Y = Point1.Top + Point1.Width \ 2

    SolP = NearestFromLine(P1, L1)

    SolPoint.Left = SolP.X - SolPoint.Width \ 2
    SolPoint.Top = SolP.Y - SolPoint.Width \ 2

    If SolP.Bool Then SolPoint.BorderColor = RGB(0, 200, 0) Else: SolPoint.BorderColor = vbRed


End Sub


