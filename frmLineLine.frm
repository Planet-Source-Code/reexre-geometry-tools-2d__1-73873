VERSION 5.00
Begin VB.Form frmLineLine 
   Caption         =   "Line Line"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TXT 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmLineLine.frx":0000
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Move Lines with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Shape Circle1 
      BorderWidth     =   2
      FillColor       =   &H0000C000&
      Height          =   255
      Left            =   5040
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   304
      X2              =   544
      Y1              =   344
      Y2              =   24
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   320
      X2              =   400
      Y1              =   272
      Y2              =   296
   End
End
Attribute VB_Name = "frmLineLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private L1         As geoLine
Private L2         As geoLine
Private C          As geoCircle
Private Inter      As geoPointVector2D
Private Pmove      As Long


Private Sub Form_Load()
    FindIntersection
    Me.Width = Me.Height * 1.618

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 4)  As Single
    Dim Dmin       As Single
    Dim I          As Long

    D(1) = DistFromPoint2(L1.P1, X, Y)
    D(2) = DistFromPoint2(L1.P2, X, Y)
    D(3) = DistFromPoint2(L2.P1, X, Y)
    D(4) = DistFromPoint2(L2.P2, X, Y)


    Dmin = 999999999999#
    For I = 1 To 4
        If D(I) < Dmin Then Dmin = D(I): Pmove = I
    Next



End Sub

Private Sub FindIntersection()


    L1 = mkLine(mkPoint(Line1.x1, Line1.y1), mkPoint(Line1.x2, Line1.y2))
    L2 = mkLine(mkPoint(Line2.x1, Line2.y1), mkPoint(Line2.x2, Line2.y2))

    Inter = IntersectOfLines(L1, L2)

    If Inter.Bool Then
        Circle1.BorderColor = RGB(0, 170, 0)
        Circle1.Left = -Circle1.Width \ 2 + Inter.X
        Circle1.Top = -Circle1.Height \ 2 + Inter.Y
    Else
        Circle1.BorderColor = vbRed
    End If

    TXT = "L1X1 " & L1.P1.X & vbCrLf
    TXT = TXT & "L1Y1 " & L1.P1.Y & vbCrLf
    TXT = TXT & "L1X2 " & L1.P2.X & vbCrLf
    TXT = TXT & "L1Y2 " & L1.P2.Y & vbCrLf

    TXT = TXT & "L2X1 " & L2.P1.X & vbCrLf
    TXT = TXT & "L2Y1 " & L2.P1.Y & vbCrLf
    TXT = TXT & "L2X2 " & L2.P2.X & vbCrLf
    TXT = TXT & "L2Y2 " & L2.P2.Y & vbCrLf

    TXT = TXT & vbCrLf

    TXT = TXT & "IX " & Inter.X & vbCrLf
    TXT = TXT & "IY " & Inter.Y



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
                Line2.x1 = X
                Line2.y1 = Y
            Case 4
                Line2.x2 = X
                Line2.y2 = Y
        End Select

        FindIntersection
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub
