VERSION 5.00
Begin VB.Form frmFILLET 
   AutoRedraw      =   -1  'True
   Caption         =   "Find Arc (of a given radius) tangent to two lines"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sRadius 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   1
      Top             =   360
      Value           =   20
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Fillet Radius"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   400
      X2              =   152
      Y1              =   224
      Y2              =   328
   End
   Begin VB.Line Line1 
      X1              =   168
      X2              =   400
      Y1              =   120
      Y2              =   240
   End
End
Attribute VB_Name = "frmFILLET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private L1         As geoLine
Private L2         As geoLine
Private Arc        As geoARC
Private Pmove      As Long







Private Sub Command1_Click()

End Sub

Private Sub Form_Load()



    L1.P1.X = Line1.x1
    L1.P1.Y = Line1.y1
    L1.P2.X = Line1.x2
    L1.P2.Y = Line1.y2
    L2.P1.X = Line2.x1
    L2.P1.Y = Line2.y1
    L2.P2.X = Line2.x2
    L2.P2.Y = Line2.y2

    DOfillet


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

        DOfillet
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub




Private Sub DOfillet()
    Dim L          As geoLine

    L1 = mkLine(mkPoint(Line1.x1, Line1.y1), mkPoint(Line1.x2, Line1.y2))
    L2 = mkLine(mkPoint(Line2.x1, Line2.y1), mkPoint(Line2.x2, Line2.y2))

    Fillet L1, L2, sRadius, Arc, False


    Line1.x1 = L1.P1.X
    Line1.y1 = L1.P1.Y
    Line1.x2 = L1.P2.X
    Line1.y2 = L1.P2.Y
    Line2.x1 = L2.P1.X
    Line2.y1 = L2.P1.Y
    Line2.x2 = L2.P2.X
    Line2.y2 = L2.P2.Y


    Me.Cls



    'MyARC Me.HDC, Arc.Circle.Center.X, Arc.Circle.Center.Y, _
     Arc.Circle.Radius, Arc.A1, Arc.A2, vbBlack
    If Arc.Circle.Center.Bool Then
        MyARC2 Me.HDC, Arc.Circle.Center.X, Arc.Circle.Center.Y, _
               Arc.Circle.Radius, Arc.x1, Arc.y1, Arc.x2, Arc.y2, vbRed

        Me.CurrentX = Arc.x1
        Me.CurrentY = Arc.y1
        Me.Print "1"
        Me.CurrentX = Arc.x2
        Me.CurrentY = Arc.y2
        Me.Print "2"

    End If


    Me.Refresh
End Sub

Private Sub sRadius_Change()
    DOfillet
End Sub

Private Sub sRadius_Scroll()
    DOfillet
End Sub
