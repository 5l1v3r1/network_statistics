VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "所有接口状态"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1440
      Top             =   1560
   End
   Begin VB.Label State 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ShowMessage
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.RtnIFMessage.Enabled = True
End Sub

Private Sub Timer1_Timer()
ShowMessage
End Sub

Private Sub ShowMessage()
Dim RtnIFMessage As AboutIFTable.MIB_IFTABLE
RtnIFMessage = RtnIFTable

Dim sSum As Long   ''总共收到(字节)
Dim rSum As Long  '''总共发送(字节)

Dim sPackSum As Long  '''总共收到(包)
Dim rPackSum As Long  ''''总共发送(包)
Dim sPackUnicast As Long   ''发送Unicast包
Dim rPackUnicast As Long   ''收到Unicast包

Dim PhyAddr As String     ''接口物理地址
Dim PortStat As String    ''接口数据

For i = 0 To RtnIFMessage.dwNumEntries - 1
sSum = sSum + RtnIFMessage.MIB_Table(i).dwInOctets
rSum = rSum + RtnIFMessage.MIB_Table(i).dwOutOctets
sPackSum = sPackSum + RtnIFMessage.MIB_Table(i).dwOutNUcastPkts + RtnIFMessage.MIB_Table(i).dwOutUcastPkts
rPackSum = rPackSum + RtnIFMessage.MIB_Table(i).dwInNUcastPkts + RtnIFMessage.MIB_Table(i).dwInUcastPkts
sPackUnicast = sPackUnicast + RtnIFMessage.MIB_Table(i).dwOutUcastPkts
rPackUnicast = rPackUnicast + RtnIFMessage.MIB_Table(i).dwInUcastPkts

PhyAddr = ""       ''清零，重组物理地址
For j = 0 To 7
If j < 7 Then             ''''必须这样，不然的话最后会多加 "_"
PhyAddr = PhyAddr & Hex(RtnIFMessage.MIB_Table(i).bPhysAddr(j)) & "-"
ElseIf j = 7 Then
PhyAddr = PhyAddr & Hex(RtnIFMessage.MIB_Table(i).bPhysAddr(j))
End If
Next
Dim IfName As String
IfName = AboutIFTable.IFMessage(i)

If PhyAddr = "0-0-0-0-0-0-0-0" Then PhyAddr = "无地址"
PortStat = PortStat & "关于接口 " & i & " 数据:" & vbCrLf & vbCrLf
PortStat = PortStat & "接口物理地址 :" & PhyAddr & vbCrLf
PortStat = PortStat & "类型 :" & AboutIFTable.TestType(RtnIFMessage.MIB_Table(i).dwType) & vbCrLf
PortStat = PortStat & "状态 :" & AboutIFTable.TestOperStatus(RtnIFMessage.MIB_Table(i).dwOperStatus) & vbCrLf
PortStat = PortStat & "传输速度 :" & Sizes(RtnIFMessage.MIB_Table(i).dwSpeed, 6) & "/S" & vbCrLf

Next
If Send = 0 Then Send = sSum
If Recv = 0 Then Recv = rSum

Dim Message As String
Message = "本机共接收: " & Sizes(sSum, 7) & vbCrLf
Message = Message + "本机共收到数据包: " & sPackSum & " (个)" & vbCrLf
Message = Message + "其中 Unicast包:" & sPackUnicast & "(个)" & "  Non-Unicast包:" & sPackSum - sPackUnicast & "(个)" & vbCrLf & vbCrLf
Message = Message + "本机共发送: " & Sizes(rSum, 7) & vbCrLf
Message = Message + "本机共发送数据包: " & rPackSum & " (个)" & vbCrLf
Message = Message + "其中 Unicast包:" & rPackUnicast & "(个)" & "  Non-Unicast包:" & rPackSum - rPackUnicast & "(个)" & vbCrLf & vbCrLf
Message = Message + PortStat

State = Message + vbCrLf

'''''    ----  下面是组装 本机的 TCP/UDP 信息
Dim PutTCPMessage As LocalLine.Rtn_TCPStat
Dim PutUDPMessage As LocalLine.Rtn_UDPStat
Dim PutICMPMessage As LocalLine.Rtn_ICMPStat

 PutTCPMessage = LocalLine.RtnTCPStat
 PutUDPMessage = LocalLine.RtnUDPStat
PutICMPMessage = LocalLine.RtnICMPStat

'''    ---   开始组装信息
Message = "本机的 TCP/UDP 信息: " + vbCrLf
Message = Message & "TCP 信息:" + vbCrLf
Message = Message & "  本机TCP连接数: " & PutTCPMessage.NumConns & vbCrLf
Message = Message & "  监听:" & PutTCPMessage.Listening & "  连接中:" & PutTCPMessage.Connecting & "  关闭:" & PutTCPMessage.Closed & "  超时:" & PutTCPMessage.TimeWaitLine & "  未知:" & PutTCPMessage.Unkhownline & vbCrLf
Message = Message & "UDP 信息:" + vbCrLf
Message = Message & "  本机UDP连接数: " & PutUDPMessage.NumAddrs & "  Recv:" & PutUDPMessage.InDatagrams & "  Send:" & PutUDPMessage.OutDatagrams & vbCrLf
Message = Message & "ICMP 信息:" + vbCrLf
Message = Message & " 本机ICMP收发信息: " & PutICMPMessage.SendSum & "(发送)" & PutICMPMessage.RecvSum & "(收到)" & vbCrLf

State = State + Message

''''调整窗体
Me.Height = State.Height + 1000
Me.Width = State.Width + 500
End Sub
