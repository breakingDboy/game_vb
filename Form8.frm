VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第八关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   4545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回首页"
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   5280
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   480
      TabIndex        =   0
      Top             =   5280
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   7
      Left            =   1980
      TabIndex        =   10
      Top             =   3780
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   6
      Left            =   780
      TabIndex        =   9
      Top             =   900
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   2700
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   4
      Left            =   900
      TabIndex        =   7
      Top             =   4440
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   3
      Left            =   3180
      TabIndex        =   6
      Top             =   4440
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3780
      TabIndex        =   5
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   900
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   435
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   2700
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   6
      Left            =   840
      Shape           =   3  'Circle
      Top             =   840
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   5
      Left            =   240
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   7
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   4
      Left            =   900
      Shape           =   3  'Circle
      Top             =   4380
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   3
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   4380
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   0
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   2
      Left            =   3780
      Shape           =   3  'Circle
      Top             =   2580
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   1
      Left            =   3180
      Shape           =   3  'Circle
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   20
      X1              =   2220
      X2              =   3960
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   73
      X1              =   3360
      X2              =   2280
      Y1              =   4680
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   74
      X1              =   1080
      X2              =   2160
      Y1              =   4680
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   30
      X1              =   3420
      X2              =   2220
      Y1              =   4620
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   40
      X1              =   1020
      X2              =   2220
      Y1              =   4680
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   62
      X1              =   3960
      X2              =   1020
      Y1              =   2820
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   51
      X1              =   420
      X2              =   3360
      Y1              =   2820
      Y2              =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   54
      X1              =   1020
      X2              =   360
      Y1              =   4680
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   32
      X1              =   3420
      X2              =   4020
      Y1              =   4680
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   21
      X1              =   4020
      X2              =   3420
      Y1              =   2880
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   65
      X1              =   360
      X2              =   960
      Y1              =   2880
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   43
      X1              =   1020
      X2              =   3420
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   50
      X1              =   420
      X2              =   2160
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   10
      Index           =   61
      X1              =   1020
      X2              =   3420
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第八关—"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1620
      TabIndex        =   2
      Top             =   5340
      Width           =   1335
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim firstpoint As Boolean
Dim index2 As Integer
Dim n As Integer
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
q = MsgBox("您确定要退出游戏？", vbYesNo + 48, "退出")
   If q = 6 Then
   Unload Form0
   Else
   Cancel = 1
   End If
End If
End Sub
Private Sub Command1_Click()
firstpoint = True
n = 0
For i = 0 To 7
Shape1(i).BackColor = &H404040
Next i
Line1(61).BorderColor = &H404040
Line1(65).BorderColor = &H404040
Line1(51).BorderColor = &H404040
Line1(62).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(50).BorderColor = &H404040
Line1(20).BorderColor = &H404040
Line1(54).BorderColor = &H404040
Line1(40).BorderColor = &H404040
Line1(30).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(74).BorderColor = &H404040
Line1(73).BorderColor = &H404040
Line1(43).BorderColor = &H404040

End Sub

Private Sub Command2_Click()
Form0.Show
Unload Me
End Sub

Private Sub Form_Load()
firstpoint = True
n = 0
Me.Height = 6630
Me.Width = 4635
Me.Left = 7980
Me.Top = 2175
End Sub

Private Sub Label1_Click(Index As Integer)
If firstpoint Then
   Shape1(Index).BackColor = &HC0C000
   firstpoint = False
   index2 = Index
Else
   Select Case tpl(Index, index2)
   Case 61, 65, 51, 62, 21, 50, 20, 54, 40, 30, 32, 74, 73, 43
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 14 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form9.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

