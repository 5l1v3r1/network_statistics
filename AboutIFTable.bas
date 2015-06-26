Attribute VB_Name = "AboutIFTable"
Private Declare Function GetIfTable Lib "iphlpapi.dll" (ByRef pIfTable As MIB_IFTABLE, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

Public Type MIB_IFROW
    wszName(0 To 511) As Byte '�ӿ����Ƶ�Unicode�ַ���������Ϊ512�ֽ�
    dwIndex As Long         '�ӿڱ��
    dwType As Long          '�ӿ����ͣ��ο�IP_ADAPTER_INFO���͵�Type��Ա
    dwMtu As Long           '����䵥Ԫ
    dwSpeed As Long         '�ӿ��ٶȣ��ֽڣ�
    dwPhysAddrLen As Long   '��bPhysAddr��õ������ַ��Ч����
    bPhysAddr(0 To 7) As Byte  '�����ַ
    dwAdminStatus As Long    '�ӿڹ���״̬
dwOperStatus As Long      '����״̬������ֵ֮һ��

    dwLastChange As Long     '����״̬���ı��ʱ��
    dwInOctets As Long       '�ܹ��յ�(�ֽ�)
    dwInUcastPkts As Long    '�ܹ��յ�(unicast��)
    dwInNUcastPkts As Long   '�ܹ��յ�(non-unicast��)�������㲥���Ͷ�㴫�Ͱ�
    dwInDiscards As Long      '�յ���������������ʹû�д���
    dwInErrors As Long       '�յ����������
    dwInUnknownProtos As Long   '�յ�����Э�鲻���������İ�����
    dwOutOctets As Long      '�ܹ�����(�ֽ�)
    dwOutUcastPkts As Long   '�ܹ�����(unicast��)
    dwOutNUcastPkts As Long  '�ܹ�����(non-unicast��)�������㲥���Ͷ�㴫�Ͱ�
    dwOutDiscards As Long    '���Ͷ�������������ʹû�д���
    dwOutErrors As Long      '���ͳ��������
    dwOutQLen As Long       '���Ͷ��г���
    dwDescrLen As Long      ' bDescr������Ч����
    bDescr(0 To 255) As Byte   '�ӿ�����
End Type

Public Type MIB_IFTABLE
    dwNumEntries As Long            '��ǰ����ӿڵ�����
    MIB_Table(9) As MIB_IFROW     'ָ��һ������MIB_IFROW���͵�ָ��
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
Sizes = Left(size, Bit) & " B"           ''''����Size��ȥ���С����̫�࣬����ֻȡ 4λ
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

''�������� ֵ ˵��''
''MIB_IF_TYPE_ETHERNET 6 ��̫��������
''MIB_IF_TYPE_TOKENRING 9 ���ƻ�������
''MIB_IF_TYPE_FDDI 15 ���˽ӿ�������
''MIB_IF_TYPE_PPP 23 �㵽��Э��������
''MIB_IF_TYPE_LOOPBACK 24 �ػ���Loopback��������
''MIB_IF_TYPE_SLIP 28 ����������(Serial Line Interface Protocol)
''MIB_IF_TYPE_OTHER ����ֵ �������͵�������

If StatusValue = 6 Then
TestType = "��̫��������"
ElseIf StatusValue = 9 Then
TestType = "���ƻ�������"
ElseIf StatusValue = 15 Then
TestType = "���˽ӿ�������"
ElseIf StatusValue = 23 Then
TestType = "�㵽��Э��������"
ElseIf StatusValue = 24 Then
TestType = "�ػ�������"
ElseIf StatusValue = 28 Then
TestType = "����������"
Else
TestType = "�������͵�����������"
End If

End Function

Public Function TestOperStatus(ByVal StatusValue As Long) As String

'�������� ֵ ˵��
'MIB_IF_OPER_STATUS_NON_OPERATIONAL 0 ��������������ֹ�����磺��ַ��ͻ
'MIB_IF_OPER_STATUS_UNREACHABLE 1 û������
'MIB_IF_OPER_STATUS_DISCONNECTED 2 ������������δ���ӣ������������ز��ź�
'MIB_IF_OPER_STATUS_CONNECTING 3 ������������������
'MIB_IF_OPER_STATUS_CONNECTED 4 ������������������Զ�̶Եȵ�
'MIB_IF_OPER_STATUS_OPERATIONAL 5 ������������Ĭ��״̬

If StatusValue = 0 Then
TestOperStatus = "��������������ֹ�����磺��ַ��ͻ"
ElseIf StatusValue = 1 Then
TestOperStatus = "û������"
ElseIf StatusValue = 2 Then
TestOperStatus = "������������δ���ӣ������������ز��ź�"
ElseIf StatusValue = 3 Then
TestOperStatus = "������������������"
ElseIf StatusValue = 4 Then
TestOperStatus = "������������������Զ�̶Եȵ�"
ElseIf StatusValue = 5 Then
TestOperStatus = "������������Ĭ��״̬"
Else
TestOperStatus = "δ֪״̬"
End If
End Function
