VERSION 5.00
Begin VB.Form frmAngleDiff 
   Caption         =   "Angles Difference"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   352
      X2              =   360
      Y1              =   184
      Y2              =   96
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   352
      X2              =   480
      Y1              =   184
      Y2              =   216
   End
   Begin VB.Label Label1 
      Caption         =   "Move Lines with Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmAngleDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private L1         As geoLine
Private L2         As geoLine

Private Pmove      As Long


Private Sub Form_Load()
    FindAngleDiff
    Me.Width = Me.Height * 1.618

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim D(1 To 4)  As Single
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

        FindAngleDiff
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub


Private Sub FindAngleDiff()
    Dim A          As Single

    L1.P1.X = Line1.x1
    L1.P1.Y = Line1.y1
    L1.P2.X = Line1.x2
    L1.P2.Y = Line1.y2

    L2.P1.X = Line2.x1
    L2.P1.Y = Line2.y1
    L2.P2.X = Line2.x2
    L2.P2.Y = Line2.y2

    UpdateLineAng L1
    UpdateLineAng L2

    A = AngleDIFF(L1.Ang, L2.Ang)


    TXT = "Clockwise" & vbCrLf
    TXT = TXT & 180 * L1.Ang / PI & vbCrLf
    TXT = TXT & 180 * L2.Ang / PI & vbCrLf & vbCrLf
    TXT = TXT & 180 * A / PI & vbCrLf

End Sub
