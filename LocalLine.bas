Attribute VB_Name = "LocalLine"
Private Declare Function GetTcpStatistics Lib "iphlpapi.dll" (ByRef pTcpStats As MIB_TCPSTATS) As Long
Private Type MIB_TCPSTATS
dwRtoAlgorithm As Long      '指定重传输(RTO：retransmission time-out)算法
    dwRtoMin As Long             ' '‘重传输超时的最小值，毫秒
    dwRtoMax As Long              '‘重传输超时的最大值，毫秒
    dwMaxConn As Long         '  ‘连接最大数目，如果为-1，则连接的最大数目是可变的
    dwActiveOpens As Long      ' ‘主动连接数目，即客户端正向服务器进行连接数目
    dwPassiveOpens As Long     '‘被动连接数目，即服务器监听连接客户端请求数目
    dwAttemptFails As Long      ' ‘尝试连接失败的次数
    dwEstabResets As Long        '‘对已建立的连接实行重设的次数
    dwCurrEstab As Long           '‘目前已建立的连接
    dwInSegs As Long               '‘收到分段数据报的数目
    dwOutSegs As Long             '‘'传输的分段数据报数目，不包括转发的数据包
    dwRetransSegs As Long          '    ‘转发的分段数据报数目
    dwInErrs As Long               ' ‘收到错误的数目
    dwOutRsts As Long          '   ‘重设标志设定后传输分段数据报数目
    dwNumConns As Long         '‘累计连接的总数
End Type

Private Declare Function GetUdpStatistics Lib "iphlpapi.dll" (pStats As MIB_UDPSTATS) As Long
Private Type MIB_UDPSTATS
    dwInDatagrams As Long   '已收到数据报数目
    dwNoPorts As Long       '因为端口号有误而丢弃的数据报数目
    dwInErrors As Long        '已收到多少错误数据报，不包括dwNoPorts中统计的数目
    dwOutDatagrams As Long  '已传输数据报数目
    dwNumAddrs As Long     'UDP监听者表中接口数目
End Type


Private Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIB_ICMP) As Long
Private Type MIBICMPSTATS
    dwMsgs As Long          '已收发多少消息
    dwErrors As Long          '已收发多少错误
    dwDestUnreachs As Long    '已收发多少"目标不可抵达"消息
    dwTimeExcds As Long            '已收发多少生存期已过消息
    dwParmProbs As Long            '已收发多少表明数据报内有错误IP信息的消息
    dwSrcQuenchs As Long           '已收发多少源结束消息
    dwRedirects As Long        '已收发多少重定向消息
    dwEchos As Long                '已收发多少ICMP响应请求
    dwEchoReps As Long             '已收发多少ICMP响应应答
    dwTimestamps As Long    '已收发多少时间戳请求
    dwTimestampReps As Long        '已收发多少时间戳响应
    dwAddrMasks As Long      '已收发多少地址掩码
    dwAddrMaskReps As Long '已收发多少地址掩码响应
End Type

Private Type MIBICMPINFO
  icmpInStats As MIBICMPSTATS   '指向MIBICMPSTATS类型，包含接收数据
  icmpOutStats As MIBICMPSTATS '指向MIBICMPSTATS类型，包含发出数据
End Type
Private Type MIB_ICMP
    stats As MIBICMPINFO    '指定MIBICMPINFO类型包含了电脑ICMP统计信息表
End Type

'' 查看TCP表
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Private Type MIB_TCPROW ' TCP连接表中一行的结构
dwState As Long ' 状态
dwLocalAddr As Long ' Local IP
dwLocalPort As Long ' Local port
dwRemoteAddr As Long ' Remote IP
dwRemotePort As Long ' Remote port
End Type

Private Type MIB_TCPTABLE
dwNum_Of_Entries As Long ' 当前 TCP连接的总数
TCP_Table(120) As MIB_TCPROW ' 预留了120行的缓冲区
End Type

'''''    ----Return to Client's type

Public Type Rtn_TCPStat
 NumConns As Long          '‘累计连接的总数
 CurrEstab As Long           '‘目前已建立的连接
Unkhownline As Long   ''
Closed As Long
Listening As Long
Connecting As Long
TimeWaitLine As Long
 End Type
 
 Public Type Rtn_UDPStat
    InDatagrams As Long   '已收到数据报数目
    OutDatagrams As Long  '已传输数据报数目
    NumAddrs As Long     'UDP监听者表中接口数目
End Type
 
  Public Type Rtn_ICMPStat
  RecvSum As Long
  SendSum As Long
 End Type
 
 ''''     ---  Return to Client's Function
 Public Function RtnTCPStat() As Rtn_TCPStat
 Dim TCPState As MIB_TCPSTATS
 GetTcpStatistics TCPState
 RtnTCPStat.NumConns = TCPState.dwNumConns
   RtnTCPStat.CurrEstab = TCPState.dwCurrEstab
    Dim TCPLineStat As MIB_TCPTABLE
    GetTcpTable TCPLineStat, Len(TCPLineStat), 0
    
    For i = 0 To TCPLineStat.dwNum_Of_Entries - 1
    Select Case TCPLineStat.TCP_Table(i).dwState
    Case 0
    RtnTCPStat.Unkhownline = RtnTCPStat.Unkhownline + 1
    Case 1
    RtnTCPStat.Closed = RtnTCPStat.Closed + 1
    Case 2
    RtnTCPStat.Listening = RtnTCPStat.Listening + 1
    Case 5
    RtnTCPStat.Connecting = RtnTCPStat.Connecting + 1
    Case 11
    RtnTCPStat.TimeWaitLine = RtnTCPStat.TimeWaitLine + 1
    End Select
    Next
 
 End Function
 
  Public Function RtnUDPStat() As Rtn_UDPStat
Dim UDPState As MIB_UDPSTATS
GetUdpStatistics UDPState
RtnUDPStat.NumAddrs = UDPState.dwNumAddrs
RtnUDPStat.InDatagrams = UDPState.dwInDatagrams
RtnUDPStat.OutDatagrams = UDPState.dwOutDatagrams
  End Function

  Public Function RtnICMPStat() As Rtn_ICMPStat
  Dim ICMPState As MIB_ICMP
  GetIcmpStatistics ICMPState
  RtnICMPStat.RecvSum = ICMPState.stats.icmpInStats.dwAddrMaskReps + ICMPState.stats.icmpInStats.dwAddrMasks + ICMPState.stats.icmpInStats.dwDestUnreachs + ICMPState.stats.icmpInStats.dwEchoReps + ICMPState.stats.icmpInStats.dwEchos + ICMPState.stats.icmpInStats.dwMsgs + ICMPState.stats.icmpInStats.dwParmProbs + ICMPState.stats.icmpInStats.dwRedirects + ICMPState.stats.icmpInStats.dwSrcQuenchs + ICMPState.stats.icmpInStats.dwTimeExcds + ICMPState.stats.icmpInStats.dwTimestampReps + ICMPState.stats.icmpInStats.dwTimestamps
  
  RtnICMPStat.SendSum = ICMPState.stats.icmpOutStats.dwAddrMaskReps + ICMPState.stats.icmpOutStats.dwAddrMasks + ICMPState.stats.icmpOutStats.dwDestUnreachs + ICMPState.stats.icmpOutStats.dwEchoReps + ICMPState.stats.icmpOutStats.dwEchos + ICMPState.stats.icmpOutStats.dwMsgs + ICMPState.stats.icmpOutStats.dwParmProbs + ICMPState.stats.icmpOutStats.dwRedirects + ICMPState.stats.icmpOutStats.dwSrcQuenchs + ICMPState.stats.icmpOutStats.dwTimeExcds + ICMPState.stats.icmpOutStats.dwTimestampReps + ICMPState.stats.icmpOutStats.dwTimestamps
  End Function

