Attribute VB_Name = "AboutIFTable"
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As MIB_IFTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Public Type MIB_IFROW
    wszName(0 To 511) As Byte '接口名称的Unicode字符串，必须为512字节
    dwIndex As Long         '接口编号
    dwType As Long          '接口类型，参看IP_ADAPTER_INFO类型的Type成员
    dwMtu As Long           '最大传输单元
    dwSpeed As Long         '接口速度（字节）
    dwPhysAddrLen As Long   '由bPhysAddr获得的物理地址有效长度
    bPhysAddr(0 To 7) As Byte  '物理地址
    dwAdminStatus As Long    '接口管理状态
dwOperStatus As Long      '操作状态，以下值之一：

    dwLastChange As Long     '操作状态最后改变的时间
    dwInOctets As Long       '总共收到(字节)
    dwInUcastPkts As Long    '总共收到(unicast包)
    dwInNUcastPkts As Long   '总共收到(non-unicast包)，包括广播包和多点传送包
    dwInDiscards As Long      '收到后丢弃包总数（即使没有错误）
    dwInErrors As Long       '收到出错包总数
    dwInUnknownProtos As Long   '收到后因协议不明而丢弃的包总数
    dwOutOctets As Long      '总共发送(字节)
    dwOutUcastPkts As Long   '总共发送(unicast包)
    dwOutNUcastPkts As Long  '总共发送(non-unicast包)，包括广播包和多点传送包
    dwOutDiscards As Long    '发送丢弃包总数（即使没有错误）
    dwOutErrors As Long      '发送出错包总数
    dwOutQLen As Long       '发送队列长度
    dwDescrLen As Long      ' bDescr部分有效长度
    bDescr(0 To 255) As Byte   '接口描述
End Type

Public Type MIB_IFTABLE
    dwNumEntries As Long            '当前网络接口的总数
    MIB_Table(9) As MIB_IFROW     '指向一个包含MIB_IFROW类型的指针
End Type

Public Function RtnIFTable() As MIB_IFTABLE
Dim rtn As MIB_IFTABLE
GetIfTable rtn, Len(rtn), 0
RtnIFTable = rtn
End Function

Public Function IFMessage(ByVal Index As Long) As String
Dim rtn As MIB_IFTABLE
GetIfTable rtn, Len(rtn), 0
Dim RtnStr As String
For i = 0 To 255
RtnStr = RtnStr & Chr(rtn.MIB_Table(Index).bDescr(i))
Next

IFMessage = Trim(RtnStr)
End Function


Public Function Sizes(ByVal size As Double, ByVal Bit As Long) As String
If size = 0 Then Sizes = 0

If size >= 1024 Then
size = size / 1024
Else
Sizes = Left(size, Bit) & " B"           ''''由于Size除去后的小数点太多，所以只取 4位
Exit Function
End If

If size >= 1024 Then
size = size / 1024
Else
Sizes = Left(size, Bit) & " KB"
Exit Function
End If

If size >= 1024 Then
size = size / 1024
Else
Sizes = Left(size, Bit) & " MB"
Exit Function
End If

If size >= 1024 Then
size = size / 1024
Else
Sizes = Left(size, 4) & " GB"
Exit Function
End If

If size >= 1024 Then
size = size / 1024
Else
Sizes = Left(size, Bit) & " TB"
Exit Function
End If

End Function

Public Function TestType(ByVal StatusValue As Long) As String

''常量名称 值 说明''
''MIB_IF_TYPE_ETHERNET 6 以太网适配器
''MIB_IF_TYPE_TOKENRING 9 令牌环适配器
''MIB_IF_TYPE_FDDI 15 光纤接口适配器
''MIB_IF_TYPE_PPP 23 点到点协议适配器
''MIB_IF_TYPE_LOOPBACK 24 回环（Loopback）适配器
''MIB_IF_TYPE_SLIP 28 串行适配器(Serial Line Interface Protocol)
''MIB_IF_TYPE_OTHER 其他值 其他类型的适配器

If StatusValue = 6 Then
TestType = "以太网适配器"
ElseIf StatusValue = 9 Then
TestType = "令牌环适配器"
ElseIf StatusValue = 15 Then
TestType = "光纤接口适配器"
ElseIf StatusValue = 23 Then
TestType = "点到点协议适配器"
ElseIf StatusValue = 24 Then
TestType = "回环适配器"
ElseIf StatusValue = 28 Then
TestType = "串行适配器"
Else
TestType = "其他类型的适配器或无"
End If

End Function

Public Function TestOperStatus(ByVal StatusValue As Long) As String

'常量名称 值 说明
'MIB_IF_OPER_STATUS_NON_OPERATIONAL 0 网络适配器被禁止，例如：地址冲突
'MIB_IF_OPER_STATUS_UNREACHABLE 1 没有连接
'MIB_IF_OPER_STATUS_DISCONNECTED 2 局域网：电缆未连接；广域网：无载波信号
'MIB_IF_OPER_STATUS_CONNECTING 3 广域网适配器连接中
'MIB_IF_OPER_STATUS_CONNECTED 4 广域网适配器连接上远程对等点
'MIB_IF_OPER_STATUS_OPERATIONAL 5 局域网适配器默认状态

If StatusValue = 0 Then
TestOperStatus = "网络适配器被禁止，例如：地址冲突"
ElseIf StatusValue = 1 Then
TestOperStatus = "没有连接"
ElseIf StatusValue = 2 Then
TestOperStatus = "局域网：电缆未连接；广域网：无载波信号"
ElseIf StatusValue = 3 Then
TestOperStatus = "广域网适配器连接中"
ElseIf StatusValue = 4 Then
TestOperStatus = "广域网适配器连接上远程对等点"
ElseIf StatusValue = 5 Then
TestOperStatus = "局域网适配器默认状态"
Else
TestOperStatus = "未知状态"
End If
End Function
