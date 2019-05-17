VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���� MyApp"
   ClientHeight    =   3465
   ClientLeft      =   3030
   ClientTop       =   2910
   ClientWidth     =   5895
   ClipControls    =   0   'False
   Icon            =   "frmRegKAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2391.605
   ScaleMode       =   0  'User
   ScaleWidth      =   5535.71
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmRegKAbout.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   345
      Left            =   4365
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "ϵͳ��Ϣ(&S)"
      Height          =   345
      Left            =   4380
      TabIndex        =   2
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "����ߣ��Խ���  Copyright (1998-1999)"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1650.311
      Y2              =   1650.311
   End
   Begin VB.Label lblDescription 
      Caption         =   $"frmRegKAbout.frx":0614
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   4725
   End
   Begin VB.Label lblTitle 
      Caption         =   "ע��ŷ�����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   2955
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1660.664
      Y2              =   1660.664
   End
   Begin VB.Label lblVersion 
      Caption         =   "�汾��"
      Height          =   225
      Left            =   4320
      TabIndex        =   6
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmRegKAbout.frx":06AC
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   135
      TabIndex        =   4
      Top             =   2640
      Width           =   4110
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

' ע�����ȫѡ��...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' ע��� ROOT ����...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode �� Null ��β���ַ���
Const REG_DWORD = 4                      ' 32-λ����

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
  Set frmAbout = Nothing
End Sub

Private Sub Form_Load()
    Me.Caption = "���� " & App.Title
    lblVersion.Caption = "�汾 " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = "ע��ŷ�����"
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��\����...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' ��ͼ��ע���õ�ϵͳ��Ϣ����·��...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' ��֤��֪ 32 λ�ļ��汾�Ĵ���
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' ���� - �ļ�δ�ҵ�...
        Else
            GoTo SysInfoErr
        End If
    ' ���� - ע����δ�ҵ�...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "��ʱϵͳ��Ϣ��Ч", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' ѭ��ָ��
    Dim rc As Long                                          ' ���ش���
    Dim hKey As Long                                        ' �򿪵�ע����ľ��
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' ע�������������
    Dim tmpVal As String                                    ' ע�������ʱ�洢��
    Dim KeyValSize As Long                                  ' ע��������Ĵ�С
    '------------------------------------------------------------
    ' �ڸ��� {HKEY_LOCAL_MACHINE...} �´�ע���
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' ��ע���
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������...
    
    tmpVal = String$(1024, 0)                             ' ��������ռ�
    KeyValSize = 1024                                       ' ��Ǳ�����С
    
    '------------------------------------------------------------
    ' ����ע���ֵ...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' ���/������ֵ
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' �������
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ����� Null ��β���ַ���...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null �ҵ������ַ�����ȡ
    Else                                                    ' WinNT ����Ҫ�� Null �����ַ���...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null δ�ҵ��� ����ȡ�ַ���
    End If
    '------------------------------------------------------------
    ' Ϊ��ת����������ֵ����..
    '------------------------------------------------------------
    Select Case KeyValType                                  ' ������������...
    Case REG_SZ                                             ' �ַ�����ע�����������
        KeyVal = tmpVal                                     ' �����ַ���ֵ
    Case REG_DWORD                                          ' ˫����ע�����������
        For i = Len(tmpVal) To 1 Step -1                    ' ת��ÿһλ
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' һ���ַ�һ���ַ��ؽ���ֵ
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' ת��˫����Ϊ�ַ�����
    End Select
    
    GetKeyValue = True                                      ' ���سɹ�
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
    Exit Function                                           ' �˳�
    
GetKeyError:      ' ������������...
    KeyVal = ""                                             ' ���÷���ֵΪ���ַ���
    GetKeyValue = False                                     ' ����ʧ��
    rc = RegCloseKey(hKey)                                  ' �ر�ע���
End Function

