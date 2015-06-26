Attribute VB_Name = "LocalInetStat"
Private Declare Function InternetGetConnectedState Lib "Wininet.dll" (ByVal Flag As Long, ByVal Reserved As Long) As Long

Public Type Rtn_InetStat
IsConnecting As Boolean
IsModenConnecting As Boolean
IsModenBusy As Boolean
IsLANConnecting As Boolean
IsProxyConnecting As Boolean
End Type

Public Function RtnInetStat() As Rtn_InetStat
RtnInetStat.IsConnecting = RtnNunBooleanA(InternetGetConnectedState(0, 0))
RtnInetStat.IsModenConnecting = RtnNunBooleanB(InternetGetConnectedState(1, 0))
RtnInetStat.IsLANConnecting = RtnNunBooleanB(InternetGetConnectedState(2, 0))
RtnInetStat.IsProxyConnecting = RtnNunBooleanA(InternetGetConnectedState(4, 0))
RtnInetStat.IsModenBusy = RtnNunBooleanA(InternetGetConnectedState(8, 0))
End Function

Private Function RtnNunBooleanA(ByVal Num As Long) As Boolean
If Num = 0 Then
RtnNunBooleanA = False
Else
RtnNunBooleanA = True
End If
End Function

Private Function RtnNunBooleanB(ByVal Num As Long) As Boolean
If Num = 0 Then
RtnNunBooleanB = True
Else
RtnNunBooleanB = False
End If
End Function
