VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "本机Internet信息"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3195
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2640
      Top             =   2040
   End
   Begin VB.Label HostName 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   11
      Top             =   1320
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本机名:"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label IP 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "本机IP:"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Proxy:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "LAN端:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label2 
      Caption         =   "Moden:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "本地网络状态:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Proxy 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   210
   End
   Begin VB.Label LAN 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   210
   End
   Begin VB.Label Moden 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   330
   End
   Begin VB.Label Connect 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   210
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   2280
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   2280
      Picture         =   "Form4.frx":0615
      Stretch         =   -1  'True
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.InetStat.Enabled = True
End Sub

Private Sub Timer1_Timer()
If LocalInetStat.RtnInetStat.IsConnecting = True Then
Image1.Visible = False
Image2.Visible = True
Else
Image1.Visible = True
Image2.Visible = False
End If
If LocalInetStat.RtnInetStat.IsConnecting = True Then
Connect.Caption = "连接中"
Label2.Visible = True: Moden.Visible = True
Label3.Visible = True: LAN.Visible = True
Label4.Visible = True: Proxy.Visible = True
Label5.Visible = True: IP.Visible = True
Me.Height = 1950
  If LocalInetStat.RtnInetStat.IsModenConnecting = True Then
    If LocalInetStat.RtnInetStat.IsModenBusy = True Then
    Moden.Caption = "Moden繁忙"
    Else
    Moden.Caption = "Moden正常运行"
    End If
 Else
  Moden.Caption = "没有用Moden连接"
 End If
 If LocalInetStat.RtnInetStat.IsLANConnecting = True Then
 LAN.Caption = "LAN端已通过"
 Else
 LAN.Caption = "没有通过LAN端"
 End If
 If LocalInetStat.RtnInetStat.IsProxyConnecting = True Then
 Proxy.Caption = "使用中"
 Else
 Proxy.Caption = "没有代理"
 End If
IP.Caption = Form3.Socket1.LocalIP
HostName.Caption = Form3.Socket1.LocalHostName
Else
Connect.Caption = "断开连接"
Label2.Visible = False: Moden.Visible = False
Label3.Visible = False: LAN.Visible = False
Label4.Visible = False: Proxy.Visible = False
Label5.Visible = False: IP.Visible = False
Me.Height = 1300
End If
End Sub
