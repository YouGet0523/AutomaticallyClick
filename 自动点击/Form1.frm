VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "�Զ������װ1.0"
   ClientHeight    =   2895
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   4800
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "���Ե�"
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
      Caption         =   "���²���"
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
      Caption         =   "��ʼ"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "λ��Y��"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "λ��X��"
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
   Begin VB.Menu �˵� 
      Caption         =   "�˵�"
      Begin VB.Menu ���๦�� 
         Caption         =   "���๦��"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ִ��ʱ�� 
         Caption         =   "ִ��ʱ��"
      End
      Begin VB.Menu ִ���ٶ� 
         Caption         =   "ִ���ٶ�"
      End
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Begin VB.Menu ʹ��˵�� 
         Caption         =   "ʹ��˵��"
      End
      Begin VB.Menu �������� 
         Caption         =   "��������"
      End
      Begin VB.Menu ��ϵ���� 
         Caption         =   "��ϵ����"
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
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 'ָ�����ʹ�þ�������ϵ����ʱ����Ļ��ˮƽ�ʹ�ֱ�����Ͼ��ȷָ��65535��65535����Ԫ
Private Const MOUSEEVENTF_MOVE = &H1 '�ƶ����
Private Const MOUSEEVENTF_LEFTDOWN = &H2 'ģ������������
Private Const MOUSEEVENTF_LEFTUP = &H4 'ģ��������̧��

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long


Private Sub Command3_Click()
If Command3.Caption = "���Ե�" Then
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
        'Label1.Caption = "���λ�� x:" & pt.x & " y:" & pt.y
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
        Label1.Caption = "��ǰ���λ�� x:" & pt.X & " y:" & pt.Y
    End If
End Sub



Private Sub Timer3_Timer()
    tm3 = tm3 + 1
    If tm3 = tm3pd Then
    
        Timer1.Enabled = False
        MsgBox "ִ����ϣ�", , "YouGet"
        
        Timer3.Enabled = False
    End If
End Sub

Private Sub ���๦��_Click()
    MsgBox "�������С���", , "YouGet"
End Sub

Private Sub ��������_Click()
    MsgBox "���ͣ�http://youget.vip", , "YouGet"
End Sub

Private Sub ��ϵ����_Click()
    MsgBox "QQ:1377351008" & vbCrLf & "΢�źţ�13262333362", , "YouGet"
End Sub

Private Sub ʹ��˵��_Click()
    MsgBox "��סҪ���λ�õ�X��Y��Ϣ��" & vbCrLf & "����д��λ��X��Y�У�" & vbCrLf & "�ڹ����������ٶȺ�ִ��ʱ�䣬" & vbCrLf & "�����ʼ���ɡ�", , "YouGet"
End Sub

Private Sub ִ��ʱ��_Click()
    
    
    Do
        tm3pd = InputBox("������ִ�еļ����1Ϊһ�룡", "YouGet����ܰ��ʾ", 30)
    
    
    Loop Until tm3pd > 0
    
    tm3 = 0
    
    
   
End Sub

Private Sub ִ���ٶ�_Click()
   
    Do
        tput = InputBox("������ִ�е��ٶȣ�1000Ϊһ�룡��ֵԽС��ִ��Խ�죨����Ϊ0����", "YouGet����ܰ��ʾ", 500)
    Loop Until tput > 0
    Timer1.Interval = tput
    
End Sub
