VERSION 5.00
Object = "{677C01A7-783C-4AB6-9711-0E22E0238BEC}#1.0#0"; "SocketMaster.ocx"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ping����"
   ClientHeight    =   4530
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   10560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   10560
   StartUpPosition =   3  '����ȱʡ
   Begin SocketMasterOCX.Socket Socket1 
      Left            =   9240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.ListBox List1 
      CausesValidation=   0   'False
      Height          =   3900
      IntegralHeight  =   0   'False
      ItemData        =   "Form3.frx":0000
      Left            =   120
      List            =   "Form3.frx":000D
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   480
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ����"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "Ҫ���Ե�IP/Web-Site"
      ToolTipText     =   "��׼����IP��ʽ: xxx.xx.xx.xx"
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu Item 
      Caption         =   " Ping���Ը߼�ѡ�� "
      Begin VB.Menu Setting 
         Caption         =   "Ping���Բ�������"
         Begin VB.Menu Loops 
            Caption         =   "����ѭ�����Դ���"
         End
         Begin VB.Menu PackSize 
            Caption         =   "���ð���С"
         End
         Begin VB.Menu TimeOut 
            Caption         =   "���ó�ʱʱ�� (��ms��)"
         End
      End
      Begin VB.Menu no 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "�������"
      End
      Begin VB.Menu ClearList 
         Caption         =   "���״̬��"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoopTotal As Long
Dim TimeOutValue As Long
Dim SendPackSize As Long

Private Sub ClearList_Click()
List1.Clear
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim HostIP As New CSocketMaster
Command1.Enabled = False
List1.AddItem ""    '''' �� 1 ��,����

Dim rtn As Long
rtn = PingTest.IsIP(Text1.Text)
 
If PingTest.IsWebSite(Text1.Text) = True Then  ''  this is Website
HostIP
ElseIf rtn = 6 Then

List1.AddItem ("�ݲ�֧��IPv6��ping����")
ElseIf rtn = 4 Then

PrintPingIPMessage Text1.Text
Else

List1.AddItem ("IP/��ַ��ʽ����")
End If
Command1.Enabled = True
End Sub

Private Sub PrintPingIPMessage(ByVal IP As String)
Dim RtnICMPMessage As PingTest.Rtn_ICMPTestMessage
For i = 1 To LoopTotal
RtnICMPMessage = PingTest.RtnICMPTestMessage(IP, TimeOut, SendPackSize)
If RtnICMPMessage.testSuccess = False Then
List1.AddItem (time & " ���ͳ�ʱ " & "״̬:" & RtnICMPMessage.State)
Else
List1.AddItem (time & " ���͸�- " & RtnICMPMessage.address & " ��ICMP��Ϣ:" & " ICMP��״̬:" & RtnICMPMessage.State & " ��С:" & RtnICMPMessage.Size & " ����ʱ��:" & RtnICMPMessage.time & " ms " & " TTL:" & RtnICMPMessage.ttl)
End If
Next
End Sub

Private Sub PrintPingWEBMessage(ByVal IP As String, ByVal Host As String)
Dim RtnICMPMessage As PingTest.Rtn_ICMPTestMessage
For i = 1 To LoopTotal
RtnICMPMessage = PingTest.RtnICMPTestMessage(IP, TimeOut, SendPackSize)
If RtnICMPMessage.testSuccess = False Then
List1.AddItem (time & " ���ͳ�ʱ " & "״̬:" & RtnICMPMessage.State)
Else
List1.AddItem (time & " ���͸�- " & RtnICMPMessage.address & "( " & Host & " )" & " ��ICMP��Ϣ:" & " ICMP��״̬:" & RtnICMPMessage.State & " ��С:" & RtnICMPMessage.Size & " ����ʱ��:" & RtnICMPMessage.time & " ms " & " TTL:" & RtnICMPMessage.ttl)
End If
Next
End Sub

Private Sub Form_Load()
''''   ��ʼ��
 LoopTotal = 4
TimeOutValue = 0
SendPackSize = 32
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.StartPing.Enabled = True
End Sub

Private Sub Loops_Click()
On Error Resume Next
Dim Value As String
Value = InputBox("������һ�� 1-15 ����ֵ")
If 1 <= Value <= 15 Then
LoopTotal = Value
End If
If Value < 1 Or Value > 15 Then
MsgBox "����һ����Ч����ֵ"
End If
End Sub

Private Sub PackSize_Click()
On Error Resume Next
Dim Value As String
Value = InputBox("������һ�� 0-1024 ����ֵ")
If 0 <= Val(Value) <= 1024 Then SendPackSize = Value
If Val(Value) > 1024 Then MsgBox "��̫����!", vbExclamation
End Sub

Private Sub Print_Click()
Open "Ping���Խ��" & ".txt" For Output As #1
Write #1, CStr(Now)
For i = 0 To List1.ListCount
Write #1, List1.List(i)
Next
Close
End Sub

Private Sub Socket1_Connect()
PrintPingWEBMessage Socket1.RemoteHostIP, Socket1.RemoteHost
Socket1.CloseSck
End Sub

Private Sub Socket1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
List1.AddItem "������վIP����"
Socket1.CloseSck
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1_Click
End If
End Sub

Private Sub TimeOut_Click()
On Error Resume Next
Dim Value As String
Value = InputBox("������һ�� 0-50000 ����ֵ")
If 0 <= Value <= 50000 Then
TimeOutValue = Value
End If
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
List1.AddItem (time & " ������ַʧ��")
End Sub
