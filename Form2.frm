VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���нӿ�״̬"
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
   StartUpPosition =   3  '����ȱʡ
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

Dim sSum As Long   ''�ܹ��յ�(�ֽ�)
Dim rSum As Long  '''�ܹ�����(�ֽ�)

Dim sPackSum As Long  '''�ܹ��յ�(��)
Dim rPackSum As Long  ''''�ܹ�����(��)
Dim sPackUnicast As Long   ''����Unicast��
Dim rPackUnicast As Long   ''�յ�Unicast��

Dim PhyAddr As String     ''�ӿ������ַ
Dim PortStat As String    ''�ӿ�����

For i = 0 To RtnIFMessage.dwNumEntries - 1
sSum = sSum + RtnIFMessage.MIB_Table(i).dwInOctets
rSum = rSum + RtnIFMessage.MIB_Table(i).dwOutOctets
sPackSum = sPackSum + RtnIFMessage.MIB_Table(i).dwOutNUcastPkts + RtnIFMessage.MIB_Table(i).dwOutUcastPkts
rPackSum = rPackSum + RtnIFMessage.MIB_Table(i).dwInNUcastPkts + RtnIFMessage.MIB_Table(i).dwInUcastPkts
sPackUnicast = sPackUnicast + RtnIFMessage.MIB_Table(i).dwOutUcastPkts
rPackUnicast = rPackUnicast + RtnIFMessage.MIB_Table(i).dwInUcastPkts

PhyAddr = ""       ''���㣬���������ַ
For j = 0 To 7
If j < 7 Then             ''''������������Ȼ�Ļ������� "_"
PhyAddr = PhyAddr & Hex(RtnIFMessage.MIB_Table(i).bPhysAddr(j)) & "-"
ElseIf j = 7 Then
PhyAddr = PhyAddr & Hex(RtnIFMessage.MIB_Table(i).bPhysAddr(j))
End If
Next
Dim IfName As String
IfName = AboutIFTable.IFMessage(i)

If PhyAddr = "0-0-0-0-0-0-0-0" Then PhyAddr = "�޵�ַ"
PortStat = PortStat & "���ڽӿ� " & i & " ����:" & vbCrLf & vbCrLf
PortStat = PortStat & "�ӿ������ַ :" & PhyAddr & vbCrLf
PortStat = PortStat & "���� :" & AboutIFTable.TestType(RtnIFMessage.MIB_Table(i).dwType) & vbCrLf
PortStat = PortStat & "״̬ :" & AboutIFTable.TestOperStatus(RtnIFMessage.MIB_Table(i).dwOperStatus) & vbCrLf
PortStat = PortStat & "�����ٶ� :" & Sizes(RtnIFMessage.MIB_Table(i).dwSpeed, 6) & "/S" & vbCrLf

Next
If Send = 0 Then Send = sSum
If Recv = 0 Then Recv = rSum

Dim Message As String
Message = "����������: " & Sizes(sSum, 7) & vbCrLf
Message = Message + "�������յ����ݰ�: " & sPackSum & " (��)" & vbCrLf
Message = Message + "���� Unicast��:" & sPackUnicast & "(��)" & "  Non-Unicast��:" & sPackSum - sPackUnicast & "(��)" & vbCrLf & vbCrLf
Message = Message + "����������: " & Sizes(rSum, 7) & vbCrLf
Message = Message + "�������������ݰ�: " & rPackSum & " (��)" & vbCrLf
Message = Message + "���� Unicast��:" & rPackUnicast & "(��)" & "  Non-Unicast��:" & rPackSum - rPackUnicast & "(��)" & vbCrLf & vbCrLf
Message = Message + PortStat

State = Message + vbCrLf

'''''    ----  ��������װ ������ TCP/UDP ��Ϣ
Dim PutTCPMessage As LocalLine.Rtn_TCPStat
Dim PutUDPMessage As LocalLine.Rtn_UDPStat
Dim PutICMPMessage As LocalLine.Rtn_ICMPStat

 PutTCPMessage = LocalLine.RtnTCPStat
 PutUDPMessage = LocalLine.RtnUDPStat
PutICMPMessage = LocalLine.RtnICMPStat

'''    ---   ��ʼ��װ��Ϣ
Message = "������ TCP/UDP ��Ϣ: " + vbCrLf
Message = Message & "TCP ��Ϣ:" + vbCrLf
Message = Message & "  ����TCP������: " & PutTCPMessage.NumConns & vbCrLf
Message = Message & "  ����:" & PutTCPMessage.Listening & "  ������:" & PutTCPMessage.Connecting & "  �ر�:" & PutTCPMessage.Closed & "  ��ʱ:" & PutTCPMessage.TimeWaitLine & "  δ֪:" & PutTCPMessage.Unkhownline & vbCrLf
Message = Message & "UDP ��Ϣ:" + vbCrLf
Message = Message & "  ����UDP������: " & PutUDPMessage.NumAddrs & "  Recv:" & PutUDPMessage.InDatagrams & "  Send:" & PutUDPMessage.OutDatagrams & vbCrLf
Message = Message & "ICMP ��Ϣ:" + vbCrLf
Message = Message & " ����ICMP�շ���Ϣ: " & PutICMPMessage.SendSum & "(����)" & PutICMPMessage.RecvSum & "(�յ�)" & vbCrLf

State = State + Message

''''��������
Me.Height = State.Height + 1000
Me.Width = State.Width + 500
End Sub
