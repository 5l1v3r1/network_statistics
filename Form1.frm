VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����ͳ��"
   ClientHeight    =   360
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   2295
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":C3E3
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6480
      Top             =   4560
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0 B"
      Height          =   180
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0 B"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1200
      Picture         =   "Form1.frx":187C6
      Top             =   90
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "Form1.frx":18B48
      Top             =   90
      Width           =   240
   End
   Begin VB.Menu item 
      Caption         =   "ѡ��"
      Visible         =   0   'False
      Begin VB.Menu RtnIFMessage 
         Caption         =   "����ӿ���Ϣ"
      End
      Begin VB.Menu StartPing 
         Caption         =   "Ping����"
      End
      Begin VB.Menu InetStat 
         Caption         =   "������Ϣ"
      End
      Begin VB.Menu No 
         Caption         =   "-"
      End
      Begin VB.Menu ResetTimerInterval 
         Caption         =   "�ı�ˢ��Ƶ��"
         Begin VB.Menu SetHigh 
            Caption         =   "�Ը�"
         End
         Begin VB.Menu SetUnderstand 
            Caption         =   "��׼"
         End
         Begin VB.Menu SetLow 
            Caption         =   "�Ե�"
         End
      End
      Begin VB.Menu End 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
 cbSize As Long
hwnd As Long
 uId As Long
uFlags As Long
ucallbackMessage As Long
 hIcon As Long
szTip As String * 32
End Type

Dim ShellStruct As NOTIFYICONDATA

Dim Recv As Long ''''����������
Dim Send As Long  ''����������



'''  -----------------  �����ǲ˵��¼�
Private Sub RtnIFMessage_Click()
Form2.Move Screen.Width - Form2.Width, Screen.Height - (Form2.Height + 2 * Form1.Height)     ''''��Ϊ�ǰ����ĵױ��������ڵĶ�������  Me.Height + 2 * Form1.Height
Form2.Show
RtnIFMessage.Enabled = False
End Sub
Private Sub InetStat_Click()
Form4.Move Screen.Width - Form2.Width - Form4.Width, Screen.Height - (Form2.Height + 2 * Form1.Height) + Form3.Height
Form4.Show
InetStat.Enabled = False
End Sub
Private Sub StartPing_Click()
If LocalInetStat.RtnInetStat.IsConnecting = True Then
Form3.Move Screen.Width - Form2.Width - Form3.Width, Screen.Height - (Form2.Height + 2 * Form1.Height)
Form3.Show
StartPing.Enabled = False
Else
MsgBox "���û�:���ڵ��Զ���,���ܲ���", vbQuestion, ""
End If
End Sub
Private Sub ResetTimerInterval_Click()
Select Case Timer1.Interval
Case 750
SetHigh.Enabled = False
SetUnderstand.Enabled = True
SetLow.Enabled = True
Case 1000
SetHigh.Enabled = True
SetUnderstand.Enabled = False
SetLow.Enabled = True
Case 1250
SetHigh.Enabled = True
SetUnderstand.Enabled = True
SetLow.Enabled = False
End Select
End Sub
Private Sub SetHigh_Click()
Timer1.Interval = 750
End Sub
Private Sub SetLow_Click()
Timer1.Interval = 1250
End Sub
Private Sub SetUnderstand_Click()
Timer1.Interval = 1000
End Sub
Private Sub End_Click()
Shell_NotifyIconA 2, ShellStruct
End
End Sub
'''  -----------------  �����ǲ˵��¼�

Private Sub Timer1_Timer()
Dim RtnTable As AboutIFTable.MIB_IFTABLE
RtnTable = RtnIFTable

Dim SendSum As Long   ''�ܹ��յ�(�ֽ�)
Dim RecvSum As Long  '''�ܹ�����(�ֽ�)

For i = 0 To RtnTable.dwNumEntries - 1
SendSum = SendSum + RtnTable.MIB_Table(i).dwInOctets
RecvSum = RecvSum + RtnTable.MIB_Table(i).dwOutOctets
Next

If Send = 0 Then Send = SendSum
If Recv = 0 Then Recv = RecvSum

Label1.Caption = AboutIFTable.Sizes(SendSum - Send, 5)
Label2.Caption = AboutIFTable.Sizes(RecvSum - Recv, 5)
Send = SendSum
Recv = RecvSum

If LocalInetStat.RtnInetStat.IsConnecting = True Then
ShellStruct.szTip = "����������" & Chr(0)
Shell_NotifyIconA 1, ShellStruct
Else
ShellStruct.szTip = "����Ͽ�" & Chr(0)
Shell_NotifyIconA 1, ShellStruct
End If
End Sub

''''       -------    �����¼�

Private Sub Form_Load()
App.TaskVisible = False
Me.Move Screen.Width - Me.Width, Screen.Height - 2 * Me.Height        ''''�������� Screen.Height - 2 * Me.Height ����Ҳ����,�����ոպ�
ShellStruct.hwnd = Me.hwnd
ShellStruct.uFlags = 2 Or 4
ShellStruct.hIcon = Me.Icon.Handle
Shell_NotifyIconA 0, ShellStruct
RemoveMenu GetSystemMenu(Me.hwnd, 0), &HF060, &H1000       ''ɾ�����ڹرռ�
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIconA 2, ShellStruct
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu Item
End If
End Sub

