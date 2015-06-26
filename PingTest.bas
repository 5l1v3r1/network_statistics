Attribute VB_Name = "PingTest"
Private Declare Function IcmpCreateFile Lib "iphlpapi.dll" () As Long
Private Declare Function IcmpSendEcho Lib "iphlpapi.dll" (ByVal ICMPHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function IcmpCloseHandle Lib "iphlpapi.dll" (ByVal ICMPHandle As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Type IP_OPTION_INFORMATION
   ttl     As Byte                    '    ������ʱ��
   Tos    As Byte                   '     ����������
   Flags As Byte                    '    ��IPͷ��־
   OptionsSize      As Byte         ' ��ѡ�����ݵĴ�С���ֽ�
   OptionsData     As Long        ' ��ָ��ѡ�����ݵ�ָ��
End Type

Private Type ICMP_ECHO_REPLY
   address  As Long                  ''���������ظ���IP��ַ
   Status  As Long          '          �������ظ���״̬���ο�����ĳ������֣�
   RoundTripTime  As Long     ' ������ʱ��RTT(����)
   DataSize  As Integer          '   ���ظ����ݴ�С(�ֽ�)
   Reserved  As Integer           '  ������
   ptrData  As Long               '   ��ָ��ظ����ݵ�ָ��
   Options  As IP_OPTION_INFORMATION '���ظ�ѡ��
   Data  As String * 250
End Type

Public Type Rtn_ICMPTestMessage
address As String
State As String     '''�ú���д
testSuccess As Boolean '''  �Ƿ�ɹ�����
time As Long
size As Long
ttl As Long
End Type

Public Function RtnICMPTestMessage(ByVal TestIP As String, ByVal TimeOut As Long, ByVal PackSize As Long) As Rtn_ICMPTestMessage
Dim ICMPHandle As Long
ICMPHandle = IcmpCreateFile    '''�������Ծ��
Dim ICMPReply As ICMP_ECHO_REPLY
Dim LongIPAdde As Long
LongIPAddr = inet_addr(TestIP)

Dim SendData As String    '''''  Ϊ�˰���С������
SendData = Space(PackSize)
IcmpSendEcho ICMPHandle, LongIPAddr, SendData, Len(SendData), 0, ICMPReply, Len(ICMPReply), TimeOut
If ICMPReply.Status = 0 Then
RtnICMPTestMessage.address = TestIP
RtnICMPTestMessage.size = ICMPReply.DataSize
RtnICMPTestMessage.State = RtnICMPState(ICMPReply.Status)
RtnICMPTestMessage.time = ICMPReply.RoundTripTime
RtnICMPTestMessage.ttl = ICMPReply.Options.ttl
RtnICMPTestMessage.testSuccess = True  '''���Գɹ�

Else
RtnICMPTestMessage.address = TestIP
RtnICMPTestMessage.State = RtnICMPState(ICMPReply.Status)
RtnICMPTestMessage.testSuccess = False  '''����ʧ��
End If

IcmpCloseHandle ICMPHandle
End Function

Private Function RtnICMPState(ByVal StateValue As Long) As String
If StateValue = 0 Then
RtnICMPState = "�ɹ�"
ElseIf StateValue = 11001 Then
RtnICMPState = "����̫С"
ElseIf StateValue = 11002 Then
RtnICMPState = "Ŀ�ĵ����粻�ܵ���"
ElseIf StateValue = 11003 Then
RtnICMPState = "Ŀ�ĵ��������ܵ���"
ElseIf StateValue = 11004 Then
RtnICMPState = "Ŀ�ĵ�Э�鲻�ܵ���"
ElseIf StateValue = 11005 Then
RtnICMPState = "Ŀ�ĵض˿ڲ��ܵ���"
ElseIf StateValue = 11006 Then
RtnICMPState = "û����Դ"
ElseIf StateValue = 11007 Then
RtnICMPState = "����ѡ��"
ElseIf StateValue = 11008 Then
RtnICMPState = "Ӳ������"
ElseIf StateValue = 11009 Then
RtnICMPState = "��Ϣ��̫��"
ElseIf StateValue = 11010 Then
RtnICMPState = "����ʱ"
ElseIf StateValue = 11011 Then
RtnICMPState = "��������"
ElseIf StateValue = 11012 Then
RtnICMPState = "����·��"
ElseIf StateValue = 11013 Then
RtnICMPState = "TTL��ֹ����"
ElseIf StateValue = 11014 Then
RtnICMPState = "TTL��ֹ������װ"
ElseIf StateValue = 11015 Then
RtnICMPState = "����������"
ElseIf StateValue = 11016 Then
RtnICMPState = "��Դ����"
ElseIf StateValue = 11017 Then
RtnICMPState = "ѡ��̫��"
ElseIf StateValue = 11018 Then
RtnICMPState = "����Ŀ�ĵ�"
ElseIf StateValue = 11032 Then
RtnICMPState = "̸��IPSEC"
ElseIf StateValue = 11050 Then
RtnICMPState = "����ʧ��"
End If
End Function


Public Function IsIP(ByVal TestString As String) As Long
Dim IP As String
IP = TestString

For i = 1 To 3    ''' �ж����Ƿ�Ϊһ����׼��IPv4    4-bit

If InStr(IP, ".") <> 0 And 0 < Len(IP) - Len(Mid(IP, InStr(IP, "."))) <= 4 Then
  If InStr(IP, ".") <> 0 And i = 3 Then
  IsIP = 4   ''' IPv4
  Exit Function
  End If
 IP = Mid(IP, InStr(IP, ".") + 1)
 Else
 Exit For
End If
Next

IP = TestString
For i = 1 To 7    '''  IPv6   8-bit
If InStr(IP, ":") <> 0 And i = 7 Then
IsIP = 6    ''' IPv6
Exit Function
ElseIf InStr(IP, ":") <> 0 And 0 < Len(Right(IP, Len(IP) - InStr(IP, ":"))) <= 2 Then
IP = Mid(IP, InStr(IP, ":") + 1)
Else
Exit For
End If
Next

IsIP = 0
End Function

Public Function IsWebSite(ByVal TestString As String) As Boolean
If InStr(TestString, "www.") <> 0 Or InStr(TestString, ".com") <> 0 _
 Or InStr(TestString, ".cn") <> 0 Or InStr(TestString, ".gov") <> 0 _
 Or InStr(TestString, ".edu") <> 0 Or InStr(TestString, ".org") <> 0 _
 Or InStr(TestString, ".net") <> 0 Then
IsWebSite = True
Else
IsWebSite = False
End If
End Function
