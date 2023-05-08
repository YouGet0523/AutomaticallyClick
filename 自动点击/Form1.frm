VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "自动点击套装1.0"
   ClientHeight    =   2895
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "测试点"
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1320
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新测试"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "位置Y："
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "位置X："
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   7095
   End
   Begin VB.Menu 菜单 
      Caption         =   "菜单"
      Begin VB.Menu 更多功能 
         Caption         =   "更多功能"
      End
   End
   Begin VB.Menu 功能 
      Caption         =   "功能"
      Begin VB.Menu 执行时间 
         Caption         =   "执行时间"
      End
      Begin VB.Menu 执行速度 
         Caption         =   "执行速度"
      End
   End
   Begin VB.Menu 帮助 
      Caption         =   "帮助"
      Begin VB.Menu 使用说明 
         Caption         =   "使用说明"
      End
      Begin VB.Menu 关于我们 
         Caption         =   "关于我们"
      End
      Begin VB.Menu 联系作者 
         Caption         =   "联系作者"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x1, y1, x2, y2, tput As Long
Dim tm3, tm3pd As Integer

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim pt As POINTAPI
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '指定鼠标使用绝对坐标系，此时，屏幕在水平和垂直方向上均匀分割成65535×65535个单元
Private Const MOUSEEVENTF_MOVE = &H1 '移动鼠标
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '模拟鼠标左键按下
Private Const MOUSEEVENTF_LEFTUP = &H4 '模拟鼠标左键抬起

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long


Private Sub Command3_Click()
If Command3.Caption = "测试点" Then
    Command3.Caption = 1
Else
    Command3.Caption = Command3.Caption + 1
End If
End Sub

Private Sub Form_Load()
    tm3 = 0
    tm3pd = 30
    tput = 500
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If GetCursorPos(pt) Then
        'Label1.Caption = "鼠标位置 x:" & pt.x & " y:" & pt.y
    End If
End Sub

Private Sub Command1_Click()
    
    x1 = Text1.Text
    y1 = Text2.Text
    tm3 = 0
    Timer3.Enabled = True
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Command3.Caption = 1
End Sub

Sub MoveMouseTo(ByVal X As Long, ByVal Y As Long)
    mouse_event MOUSEEVENTF_MOVE, X, Y, 0, 0
End Sub

Private Sub Timer1_Timer()
    
   
    
    SetCursorPos x1, y1
     
    
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
   
End Sub

Private Sub Timer2_Timer()
     If GetCursorPos(pt) Then
        Label1.Caption = "当前鼠标位置 x:" & pt.X & " y:" & pt.Y
    End If
End Sub



Private Sub Timer3_Timer()
    tm3 = tm3 + 1
    If tm3 = tm3pd Then
    
        Timer1.Enabled = False
        MsgBox "执行完毕！", , "YouGet"
        
        Timer3.Enabled = False
    End If
End Sub

Private Sub 更多功能_Click()
    MsgBox "待开发中……", , "YouGet"
End Sub

Private Sub 关于我们_Click()
    MsgBox "博客：http://youget.vip", , "YouGet"
End Sub

Private Sub 联系作者_Click()
    MsgBox "QQ:1377351008" & vbCrLf & "微信号：13262333362", , "YouGet"
End Sub

Private Sub 使用说明_Click()
    MsgBox "记住要点击位置的X，Y信息，" & vbCrLf & "并填写到位置X，Y中，" & vbCrLf & "在功能栏设置速度和执行时间，" & vbCrLf & "点击开始即可。", , "YouGet"
End Sub

Private Sub 执行时间_Click()
    
    
    Do
        tm3pd = InputBox("请输入执行的间隔，1为一秒！", "YouGet的温馨提示", 30)
    
    
    Loop Until tm3pd > 0
    
    tm3 = 0
    
    
   
End Sub

Private Sub 执行速度_Click()
   
    Do
        tput = InputBox("请输入执行的速度，1000为一秒！数值越小，执行越快（不能为0）！", "YouGet的温馨提示", 500)
    Loop Until tput > 0
    Timer1.Interval = tput
    
End Sub
