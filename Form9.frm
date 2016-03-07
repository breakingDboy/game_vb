VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "第九关"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form9"
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
      Top             =   5340
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重玩"
      Height          =   435
      Left            =   600
      TabIndex        =   0
      Top             =   5340
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   10
      Left            =   480
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   9
      Left            =   480
      TabIndex        =   12
      Top             =   1860
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   8
      Left            =   480
      TabIndex        =   11
      Top             =   3180
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   7
      Left            =   480
      TabIndex        =   10
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   6
      Left            =   1980
      TabIndex        =   9
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   5
      Left            =   3420
      TabIndex        =   8
      Top             =   4200
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   4
      Left            =   3420
      TabIndex        =   7
      Top             =   3180
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   3
      Left            =   3420
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   2
      Left            =   3420
      TabIndex        =   5
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   1980
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   3060
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   10
      Left            =   480
      Shape           =   3  'Circle
      Top             =   420
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   8
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   9
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   5
      Left            =   3420
      Shape           =   3  'Circle
      Top             =   4140
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   7
      Left            =   480
      Shape           =   3  'Circle
      Top             =   4140
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   6
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   4140
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   4
      Left            =   3420
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   3
      Left            =   3420
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   435
      Index           =   2
      Left            =   3420
      Shape           =   3  'Circle
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   615
      Index           =   0
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   3060
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000B&
      BorderStyle     =   0  'Transparent
      Height          =   555
      Index           =   1
      Left            =   1980
      Shape           =   3  'Circle
      Top             =   420
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   32
      X1              =   3600
      X2              =   3600
      Y1              =   720
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   109
      X1              =   660
      X2              =   660
      Y1              =   720
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   65
      X1              =   3600
      X2              =   2160
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   40
      X1              =   3540
      X2              =   2100
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   21
      X1              =   3540
      X2              =   2100
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   86
      X1              =   2100
      X2              =   660
      Y1              =   4440
      Y2              =   3420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   50
      X1              =   3600
      X2              =   2160
      Y1              =   4440
      Y2              =   3420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   64
      X1              =   2100
      X2              =   3480
      Y1              =   4440
      Y2              =   3420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   70
      X1              =   720
      X2              =   2100
      Y1              =   4380
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   54
      X1              =   3600
      X2              =   3600
      Y1              =   4440
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   87
      X1              =   660
      X2              =   660
      Y1              =   4440
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   31
      X1              =   2160
      X2              =   3480
      Y1              =   720
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   90
      X1              =   780
      X2              =   2100
      Y1              =   2100
      Y2              =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   30
      X1              =   3540
      X2              =   2160
      Y1              =   1980
      Y2              =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   91
      X1              =   2100
      X2              =   720
      Y1              =   720
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   43
      X1              =   3600
      X2              =   3600
      Y1              =   2040
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   98
      X1              =   660
      X2              =   660
      Y1              =   2100
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   101
      X1              =   2100
      X2              =   660
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   76
      X1              =   2040
      X2              =   660
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   2  'Dash
      BorderWidth     =   10
      Index           =   80
      X1              =   2040
      X2              =   660
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—第九关—"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
End
Attribute VB_Name = "Form9"
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
For i = 0 To 10
Shape1(i).BackColor = &H404040
Next i
Line1(101).BorderColor = &H404040
Line1(21).BorderColor = &H404040
Line1(109).BorderColor = &H404040
Line1(91).BorderColor = &H404040
Line1(31).BorderColor = &H404040
Line1(32).BorderColor = &H404040
Line1(98).BorderColor = &H404040
Line1(90).BorderColor = &H404040
Line1(30).BorderColor = &H404040
Line1(43).BorderColor = &H404040
Line1(80).BorderColor = &H404040
Line1(40).BorderColor = &H404040
Line1(87).BorderColor = &H404040
Line1(54).BorderColor = &H404040
Line1(86).BorderColor = &H404040
Line1(70).BorderColor = &H404040
Line1(64).BorderColor = &H404040
Line1(50).BorderColor = &H404040
Line1(76).BorderColor = &H404040
Line1(65).BorderColor = &H404040

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
   Case 101, 21, 109, 91, 31, 32, 98, 90, 30, 43, 80, 40, 87, 86, 70, 64, 50, 54, 76, 65
      If Not Line1(tpl(index2, Index)).BorderColor = &HC0C000 Then
         Line1(tpl(index2, Index)).BorderColor = &HC0C000
         Shape1(Index).BackColor = &HC0C000
         index2 = Index
         n = n + 1
      Else
         Exit Sub
      End If
   End Select
   If n = 20 Then
    ns = MsgBox("...成功...按下yes进入下一关；按下no回到首页", vbYesNoCancel, "恭喜你闯关成功")
    Select Case ns
    Case 6
      Form10.Show
      Unload Me
    Case 7
      Form0.Show
      Unload Me
    End Select
   End If
End If
End Sub

