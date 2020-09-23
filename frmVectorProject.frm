VERSION 5.00
Begin VB.Form frmVectorProject 
   Caption         =   "VectorProject"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Line Projection 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   392
      X2              =   440
      Y1              =   168
      Y2              =   112
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   392
      X2              =   320
      Y1              =   168
      Y2              =   256
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   392
      X2              =   512
      Y1              =   168
      Y2              =   184
   End
   Begin VB.Label Label1 
      Caption         =   "Move Vectors with Mouse (Green Vector =  Result of Blue vector Projected to Black Vector)"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmVectorProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private L1 As geoLine
'Private L2 As geoLine
'Private C As geoCircle
Private Inter      As geoPointVector2D
Private Pmove      As Long

Private P1         As geoPointVector2D
Private P2         As geoPointVector2D
Private Pproj      As geoPointVector2D


Private Sub Form_Load()
    FindProj
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 2)  As Single
    Dim Dmin       As Single
    Dim I          As Long

    D(1) = DistFromPoint2(mkPoint(Line1.x2, Line1.y2), X, Y)
    D(2) = DistFromPoint2(mkPoint(Line2.x2, Line2.y2), X, Y)


    Dmin = 999999999999#
    For I = 1 To 2
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Pmove <> 0 Then
        Select Case Pmove
            Case 1
                Line1.x2 = X
                Line1.y2 = Y
            Case 2
                Line2.x2 = X
                Line2.y2 = Y
        End Select

        FindProj
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindProj()
    P1.X = Line1.x2 - Line1.x1
    P1.Y = Line1.y2 - Line1.y1
    P2.X = Line2.x2 - Line2.x1
    P2.Y = Line2.y2 - Line2.y1


    Pproj = VectorProject(P1, P2)

    Projection.x2 = Projection.x1 + Pproj.X
    Projection.y2 = Projection.y1 + Pproj.Y



End Sub
