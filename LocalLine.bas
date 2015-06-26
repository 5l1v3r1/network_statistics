Attribute VB_Name = "LocalLine"
Private Declare Function GetTcpStatistics Lib "iphlpapi.dll" (ByRef pTcpStats As MIB_TCPSTATS) As Long
Private Type MIB_TCPSTATS
dwRtoAlgorithm As Long      'ָ���ش���(RTO��retransmission time-out)�㷨
    dwRtoMin As Long             ' '���ش��䳬ʱ����Сֵ������
    dwRtoMax As Long              '���ش��䳬ʱ�����ֵ������
    dwMaxConn As Long         '  �����������Ŀ�����Ϊ-1�������ӵ������Ŀ�ǿɱ��
    dwActiveOpens As Long      ' ������������Ŀ�����ͻ����������������������Ŀ
    dwPassiveOpens As Long     '������������Ŀ�����������������ӿͻ���������Ŀ
    dwAttemptFails As Long      ' ����������ʧ�ܵĴ���
    dwEstabResets As Long        '�����ѽ���������ʵ������Ĵ���
    dwCurrEstab As Long           '��Ŀǰ�ѽ���������
    dwInSegs As Long               '���յ��ֶ����ݱ�����Ŀ
    dwOutSegs As Long             '��'����ķֶ����ݱ���Ŀ��������ת�������ݰ�
    dwRetransSegs As Long          '    ��ת���ķֶ����ݱ���Ŀ
    dwInErrs As Long               ' ���յ��������Ŀ
    dwOutRsts As Long          '   �������־�趨����ֶ����ݱ���Ŀ
    dwNumConns As Long         '���ۼ����ӵ�����
End Type

Private Declare Function GetUdpStatistics Lib "iphlpapi.dll" (pStats As MIB_UDPSTATS) As Long
Private Type MIB_UDPSTATS
    dwInDatagrams As Long   '���յ����ݱ���Ŀ
    dwNoPorts As Long       '��Ϊ�˿ں���������������ݱ���Ŀ
    dwInErrors As Long        '���յ����ٴ������ݱ���������dwNoPorts��ͳ�Ƶ���Ŀ
    dwOutDatagrams As Long  '�Ѵ������ݱ���Ŀ
    dwNumAddrs As Long     'UDP�����߱��нӿ���Ŀ
End Type


Private Declare Function GetIcmpStatistics Lib "iphlpapi.dll" (pStats As MIB_ICMP) As Long
Private Type MIBICMPSTATS
    dwMsgs As Long          '���շ�������Ϣ
    dwErrors As Long          '���շ����ٴ���
    dwDestUnreachs As Long    '���շ�����"Ŀ�겻�ɵִ�"��Ϣ
    dwTimeExcds As Long            '���շ������������ѹ���Ϣ
    dwParmProbs As Long            '���շ����ٱ������ݱ����д���IP��Ϣ����Ϣ
    dwSrcQuenchs As Long           '���շ�����Դ������Ϣ
    dwRedirects As Long        '���շ������ض�����Ϣ
    dwEchos As Long                '���շ�����ICMP��Ӧ����
    dwEchoReps As Long             '���շ�����ICMP��ӦӦ��
    dwTimestamps As Long    '���շ�����ʱ�������
    dwTimestampReps As Long        '���շ�����ʱ�����Ӧ
    dwAddrMasks As Long      '���շ����ٵ�ַ����
    dwAddrMaskReps As Long '���շ����ٵ�ַ������Ӧ
End Type

Private Type MIBICMPINFO
  icmpInStats As MIBICMPSTATS   'ָ��MIBICMPSTATS���ͣ�������������
  icmpOutStats As MIBICMPSTATS 'ָ��MIBICMPSTATS���ͣ�������������
End Type
Private Type MIB_ICMP
    stats As MIBICMPINFO    'ָ��MIBICMPINFO���Ͱ����˵���ICMPͳ����Ϣ��
End Type

'' �鿴TCP��
Private Declare Function GetTcpTable Lib "iphlpapi.dll" (ByRef pTcpTable As MIB_TCPTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Private Type MIB_TCPROW ' TCP���ӱ���һ�еĽṹ
dwState As Long ' ״̬
dwLocalAddr As Long ' Local IP
dwLocalPort As Long ' Local port
dwRemoteAddr As Long ' Remote IP
dwRemotePort As Long ' Remote port
End Type

Private Type MIB_TCPTABLE
dwNum_Of_Entries As Long ' ��ǰ TCP���ӵ�����
TCP_Table(120) As MIB_TCPROW ' Ԥ����120�еĻ�����
End Type

'''''    ----Return to Client's type

Public Type Rtn_TCPStat
 NumConns As Long          '���ۼ����ӵ�����
 CurrEstab As Long           '��Ŀǰ�ѽ���������
Unkhownline As Long   ''
Closed As Long
Listening As Long
Connecting As Long
TimeWaitLine As Long
 End Type
 
 Public Type Rtn_UDPStat
    InDatagrams As Long   '���յ����ݱ���Ŀ
    OutDatagrams As Long  '�Ѵ������ݱ���Ŀ
    NumAddrs As Long     'UDP�����߱��нӿ���Ŀ
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

