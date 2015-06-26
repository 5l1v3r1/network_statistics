Attribute VB_Name = "PingTest"
Private Declare Function IcmpCreateFile Lib "iphlpapi.dll" () As Long
Private Declare Function IcmpSendEcho Lib "iphlpapi.dll" (ByVal ICMPHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function IcmpCloseHandle Lib "iphlpapi.dll" (ByVal ICMPHandle As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long

Private Type IP_OPTION_INFORMATION
   ttl     As Byte                    '    ‘生存时间
   Tos    As Byte                   '     ‘服务类型
   Flags As Byte                    '    ‘IP头标志
   OptionsSize      As Byte         ' ‘选项数据的大小，字节
   OptionsData     As Long        ' ‘指向选项数据的指针
End Type

Private Type ICMP_ECHO_REPLY
   address  As Long                  ''‘包含正回复的IP地址
   Status  As Long          '          ‘包含回复的状态（参看后面的常量部分）
   RoundTripTime  As Long     ' ‘往返时间RTT(毫秒)
   DataSize  As Integer          '   ‘回复数据大小(字节)
   Reserved  As Integer           '  ‘保留
   ptrData  As Long               '   ‘指向回复数据的指针
   Options  As IP_OPTION_INFORMATION '‘回复选项
   Data  As String * 250
End Type

Public Type Rtn_ICMPTestMessage
address As String
State As String     '''用函数写
testSuccess As Boolean '''  是否成功测试
time As Long
size As Long
ttl As Long
End Type

Public Function RtnICMPTestMessage(ByVal TestIP As String, ByVal TimeOut As Long, ByVal PackSize As Long) As Rtn_ICMPTestMessage
Dim ICMPHandle As Long
ICMPHandle = IcmpCreateFile    '''创建测试句柄
Dim ICMPReply As ICMP_ECHO_REPLY
Dim LongIPAdde As Long
LongIPAddr = inet_addr(TestIP)

Dim SendData As String    '''''  为了包大小而设置
SendData = Space(PackSize)
IcmpSendEcho ICMPHandle, LongIPAddr, SendData, Len(SendData), 0, ICMPReply, Len(ICMPReply), TimeOut
If ICMPReply.Status = 0 Then
RtnICMPTestMessage.address = TestIP
RtnICMPTestMessage.size = ICMPReply.DataSize
RtnICMPTestMessage.State = RtnICMPState(ICMPReply.Status)
RtnICMPTestMessage.time = ICMPReply.RoundTripTime
RtnICMPTestMessage.ttl = ICMPReply.Options.ttl
RtnICMPTestMessage.testSuccess = True  '''测试成功

Else
RtnICMPTestMessage.address = TestIP
RtnICMPTestMessage.State = RtnICMPState(ICMPReply.Status)
RtnICMPTestMessage.testSuccess = False  '''测试失败
End If

IcmpCloseHandle ICMPHandle
End Function

Private Function RtnICMPState(ByVal StateValue As Long) As String
If StateValue = 0 Then
RtnICMPState = "成功"
ElseIf StateValue = 11001 Then
RtnICMPState = "缓存太小"
ElseIf StateValue = 11002 Then
RtnICMPState = "目的地网络不能到达"
ElseIf StateValue = 11003 Then
RtnICMPState = "目的地主机不能到达"
ElseIf StateValue = 11004 Then
RtnICMPState = "目的地协议不能到达"
ElseIf StateValue = 11005 Then
RtnICMPState = "目的地端口不能到达"
ElseIf StateValue = 11006 Then
RtnICMPState = "没有资源"
ElseIf StateValue = 11007 Then
RtnICMPState = "错误选项"
ElseIf StateValue = 11008 Then
RtnICMPState = "硬件错误"
ElseIf StateValue = 11009 Then
RtnICMPState = "信息包太大"
ElseIf StateValue = 11010 Then
RtnICMPState = "请求超时"
ElseIf StateValue = 11011 Then
RtnICMPState = "错误请求"
ElseIf StateValue = 11012 Then
RtnICMPState = "错误路由"
ElseIf StateValue = 11013 Then
RtnICMPState = "TTL终止传输"
ElseIf StateValue = 11014 Then
RtnICMPState = "TTL终止重新组装"
ElseIf StateValue = 11015 Then
RtnICMPState = "参数有问题"
ElseIf StateValue = 11016 Then
RtnICMPState = "资源结束"
ElseIf StateValue = 11017 Then
RtnICMPState = "选项太大"
ElseIf StateValue = 11018 Then
RtnICMPState = "错误目的地"
ElseIf StateValue = 11032 Then
RtnICMPState = "谈判IPSEC"
ElseIf StateValue = 11050 Then
RtnICMPState = "常规失败"
End If
End Function


Public Function IsIP(ByVal TestString As String) As Long
Dim IP As String
IP = TestString

For i = 1 To 3    ''' 判断这是否为一个标准的IPv4    4-bit

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
