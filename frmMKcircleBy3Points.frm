VERSION 5.00
Begin VB.Form frmMKcircleBy3Points 
   Caption         =   "Make Circle by 3 Points"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   628
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape resCircle 
      BorderWidth     =   2
      Height          =   615
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Point3 
      Height          =   255
      Left            =   4440
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   255
   End
   Begin VB.Shape Point2 
      Height          =   255
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   255
   End
   Begin VB.Shape Point1 
      Height          =   255
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Move Points with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMKcircleBy3Points"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private P1         As geoPointVector2D
Private P2         As geoPointVector2D
Private P3         As geoPointVector2D

Private C          As geoCircle

Private Inter      As geoPointVector2D
Private Pmove      As Long



Private Sub Form_Load()
    Dim R          As Long
    R = Point1.Width \ 2
    P1 = mkPoint(Point1.Left + R, Point1.Top + R)
    P2 = mkPoint(Point2.Left + R, Point2.Top + R)
    P3 = mkPoint(Point3.Left + R, Point3.Top + R)

    CreateCircle
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 3)  As Single
    Dim Dmin       As Single
    Dim I          As Long
    D(1) = DistFromPoint2(mkPoint(Point1.Left, Point1.Top), X, Y)
    D(2) = DistFromPoint2(mkPoint(Point2.Left, Point2.Top), X, Y)
    D(3) = DistFromPoint2(mkPoint(Point3.Left, Point3.Top), X, Y)

    Dmin = 999999999999#
    For I = 1 To 3
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pmove <> 0 Then
        Select Case Pmove
            Case 1
                Point1.Left = X - Point1.Width \ 2
                Point1.Top = Y - Point1.Height \ 2
                P1.X = X
                P1.Y = Y
            Case 2
                Point2.Left = X - Point2.Width \ 2
                Point2.Top = Y - Point2.Height \ 2
                P2.X = X
                P2.Y = Y
            Case 3
                Point3.Left = X - Point3.Width \ 2
                Point3.Top = Y - Point3.Height \ 2
                P3.X = X
                P3.Y = Y

        End Select

        CreateCircle
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub CreateCircle()

    C = mkCircle3Points(P1, P2, P3)


    resCircle.Width = C.Radius * 2
    resCircle.Height = C.Radius * 2


    resCircle.Left = C.Center.X - C.Radius
    resCircle.Top = C.Center.Y - C.Radius


End Sub


