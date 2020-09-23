VERSION 5.00
Begin VB.Form frmVectorReflect 
   Caption         =   "Vector Reflect"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT 
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Move Vectors with Mouse (Green Vector =  Result of Blue vector Reflected to Black Vector)"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   384
      X2              =   504
      Y1              =   168
      Y2              =   184
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   384
      X2              =   312
      Y1              =   168
      Y2              =   256
   End
   Begin VB.Line vReflect 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   384
      X2              =   432
      Y1              =   168
      Y2              =   112
   End
End
Attribute VB_Name = "frmVectorReflect"
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
Private PReflect   As geoPointVector2D


Private Sub Form_Load()
    FindReflection
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

        FindReflection
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pmove = 0
End Sub

Private Sub FindReflection()

    P1.X = Line1.x2 - Line1.x1
    P1.Y = Line1.y2 - Line1.y1
    P2.X = Line2.x2 - Line2.x1
    P2.Y = Line2.y2 - Line2.y1

    PReflect = VectorReflect(P2, P1)

    vReflect.x2 = vReflect.x1 + PReflect.X
    vReflect.y2 = vReflect.y1 + PReflect.Y



End Sub

