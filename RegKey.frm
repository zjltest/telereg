VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{A874630D-EE5B-11D2-B76F-C603E977A428}#2.0#0"; "REGKEY.OCX"
Begin VB.Form RegKey 
   Caption         =   "ע��ŷ�����"
   ClientHeight    =   5310
   ClientLeft      =   2010
   ClientTop       =   2025
   ClientWidth     =   7935
   Icon            =   "RegKey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "RegKey.frx":030A
   ScaleHeight     =   5310
   ScaleWidth      =   7935
   Begin VB.Frame Frame1 
      Caption         =   "��ע��"
      Height          =   5175
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   7695
      Begin CtlRegKey.RegKey RegKey1 
         Height          =   975
         Left            =   600
         TabIndex        =   40
         Top             =   3120
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   1720
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RegCode         =   "88888888"
         RegCode         =   88888888
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��"
         Height          =   375
         Left            =   5880
         TabIndex        =   34
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblmyweb 
         Caption         =   "Web:danlihua.yeah.net"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   5400
         MouseIcon       =   "RegKey.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "�ʱࣺ118000            �绰��0415-2120031"
         Height          =   255
         Left            =   840
         TabIndex        =   33
         Top             =   4680
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "��ַ������ʡ�����ж�����85-310#  �Խ��� ��"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label Label6 
         Caption         =   $"RegKey.frx":0B8E
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   7335
      End
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "RegKey.frx":0D22
      Height          =   330
      Left            =   1080
      TabIndex        =   26
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   582
      _Version        =   393216
      ListField       =   "���к�"
      Text            =   ""
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "H:\Program Files\VB\AutoReg\RegKey.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "�û���Ϣ��"
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "�û���Ϣ��"
      Height          =   3735
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   6975
      Begin VB.TextBox txtName 
         BackColor       =   &H80000016&
         Height          =   270
         Left            =   1320
         TabIndex        =   27
         Text            =   "DEMO"
         Top             =   310
         Width           =   735
      End
      Begin VB.TextBox txtFax 
         DataField       =   "�������"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         TabIndex        =   25
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox txtBZ 
         DataField       =   "��ע"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         TabIndex        =   23
         Top             =   3240
         Width           =   5415
      End
      Begin VB.TextBox txtWeb 
         DataField       =   "��ַ"
         DataSource      =   "Data1"
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   1320
         MouseIcon       =   "RegKey.frx":0D32
         MousePointer    =   99  'Custom
         TabIndex        =   21
         ToolTipText     =   "˫����Webҳ"
         Top             =   2880
         Width           =   5415
      End
      Begin VB.TextBox txtAppN 
         BackColor       =   &H80000016&
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   3600
         PasswordChar    =   "*"
         TabIndex        =   19
         Text            =   "123456789"
         Top             =   680
         Width           =   980
      End
      Begin VB.TextBox txtUserN 
         DataField       =   "�û���"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   2160
         TabIndex        =   14
         Text            =   "12345-67890-12345-67890"
         Top             =   320
         Width           =   2415
      End
      Begin VB.TextBox TxtEW 
         DataField       =   "��������"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         MouseIcon       =   "RegKey.frx":1174
         MousePointer    =   99  'Custom
         TabIndex        =   11
         ToolTipText     =   "˫������E-Mail"
         Top             =   2520
         Width           =   5415
      End
      Begin VB.TextBox TxtTF 
         DataField       =   "�绰����"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         TabIndex        =   10
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox TxtNM 
         DataField       =   "�û���"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   5415
      End
      Begin VB.TextBox TxtAR 
         DataField       =   "�û���ַ"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   5415
      End
      Begin VB.TextBox TxtPC 
         DataField       =   "��������"
         DataSource      =   "Data1"
         Height          =   270
         Left            =   5640
         TabIndex        =   7
         Top             =   680
         Width           =   1035
      End
      Begin VB.Label lblDate 
         Caption         =   "Label3"
         DataField       =   "ע����"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5640
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "ע���գ�"
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         ToolTipText     =   "˫������ע����"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         Caption         =   "������룺"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         ToolTipText     =   "˫�����Ҵ������"
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblBZ 
         Alignment       =   1  'Right Justify
         Caption         =   "��ע��"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblWeb 
         Alignment       =   1  'Right Justify
         Caption         =   "�û���ַ��"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "˫�������û���ַ"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblAppN 
         Alignment       =   1  'Right Justify
         Caption         =   "����ţ�"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2160
         Y1              =   440
         Y2              =   440
      End
      Begin VB.Label lblRegNdisplay 
         Caption         =   "0"
         DataField       =   "ע���"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblRegN 
         Alignment       =   1  'Right Justify
         Caption         =   "ע��ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         ToolTipText     =   "˫������ע���"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblUserN 
         Alignment       =   1  'Right Justify
         Caption         =   "�û��ţ�"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "˫�������û���"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblEW 
         Alignment       =   1  'Right Justify
         Caption         =   "�������䣺"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "˫�����ҵ�������"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblTF 
         Alignment       =   1  'Right Justify
         Caption         =   "�绰���룺"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "˫�����ҵ绰����"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblNM 
         Alignment       =   1  'Right Justify
         Caption         =   "�û����ƣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "˫�������û�����"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblAR 
         Alignment       =   1  'Right Justify
         Caption         =   "�û���ַ��"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         ToolTipText     =   "˫�������û���ַ"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblPC 
         Alignment       =   1  'Right Justify
         Caption         =   "�������룺"
         Height          =   255
         Left            =   4680
         TabIndex        =   2
         ToolTipText     =   "˫��������������"
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ӡ(&P)"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���(&O)"
      Height          =   375
      Left            =   5160
      TabIndex        =   37
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   6480
      TabIndex        =   38
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblAppName 
      Caption         =   "Label1"
      DataField       =   "������"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   600
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblXLH 
      Alignment       =   1  'Right Justify
      Caption         =   "���кţ�"
      Height          =   255
      Left            =   4920
      TabIndex        =   35
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblNumber 
      DataField       =   "���к�"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   5280
      TabIndex        =   17
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblSerch 
      Caption         =   "���ң�"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu File 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Exit 
         Caption         =   "�˳�(&Q)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Set 
      Caption         =   "����(&S)"
      Begin VB.Menu Edit 
         Caption         =   "�༭������Ϣ��"
      End
   End
   Begin VB.Menu About 
      Caption         =   "����(&A)"
      Begin VB.Menu Reginfo 
         Caption         =   "ע����Ϣ(&R)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu CopyRinfo 
         Caption         =   "��Ȩ��Ϣ(&C)"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "RegKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub cmdOK_Click()
   Frame1.Visible = False
End Sub

Private Sub Command1_Click()
  Printer.EndDoc
End Sub

Private Sub Command2_Click()
   Printer.Print "         ���кţ�"; lblNumber.Caption
   Printer.Print "         �û��ţ�"; txtName.Text; "-"; txtUserN.Text
   Printer.Print "         ע��ţ�"; lblRegNdisplay.Caption
   Printer.Print "         ע���գ�"; lblDate.Caption
   Printer.Print "       �������룺"; TxtPC.Text
   Printer.Print "           ��ַ��"; TxtAR.Text
   Printer.Print "           ���ƣ�"; TxtNM.Text
   Printer.Print "           �绰��"; TxtTF.Text
   Printer.Print "           ���棺"; txtFax.Text
   Printer.Print "       �������䣺"; TxtEW.Text
   Printer.Print "           ��ַ��"; txtWeb.Text
   Printer.Print "           ��ע��"; txtBZ.Text
   Printer.Print "      ____________________________________________________________"
   Printer.Print " "
  ' Printer.EndDoc
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub
Private Sub CopyRinfo_Click()
   frmAbout.Show vbModal
End Sub

Private Sub Data1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   lblSerch.Caption = "�������кţ�"
   DBCombo1.ListField = lblNumber.DataField
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
   Data1.Caption = "���к�" & lblNumber.Caption
End Sub

Private Sub DBCombo1_Change()
   Data1.Recordset.Bookmark = DBCombo1.SelectedItem
End Sub

Private Sub DBCombo1_Click(Area As Integer)
   Data1.Recordset.Bookmark = DBCombo1.SelectedItem
End Sub

Private Sub DBCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
   Data1.Recordset.Bookmark = DBCombo1.SelectedItem
End Sub
Private Sub Edit_Click()
   IfEdit = True
   frmDL.Show vbModal, Me
End Sub

Private Sub Exit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  Frame1.Visible = False
  Data1.DatabaseName = App.Path & "\RegKey.mdb"
  frmDL.Show vbModal, Me
  Data1.RecordSource = DLH
  txtName.Text = InputAppname
  txtAppN.Text = InputAppcode
  RegKey1.AppName = InputAppname
  RegKey1.AppCode = InputAppcode
  lblSerch.Caption = "�������кţ�"
  DBCombo1.ListField = lblNumber.DataField
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form = Nothing
End Sub


Private Sub Label2_Click()
 lblSerch.Caption = "����" & Label2.Caption
   DBCombo1.ListField = lblDate.DataField
End Sub

Private Sub lblAR_Click()
  lblSerch.Caption = "����" & lblAR.Caption
   DBCombo1.ListField = TxtAR.DataField
End Sub

Private Sub lblBZ_Click()
lblSerch.Caption = "����" & lblBZ.Caption
   DBCombo1.ListField = txtBZ.DataField
End Sub

Private Sub lblEW_Click()
  lblSerch.Caption = "����" & lblEW.Caption
   DBCombo1.ListField = TxtEW.DataField
End Sub

Private Sub lblFax_Click()

lblSerch.Caption = "����" & lblFax.Caption
   DBCombo1.ListField = txtFax.DataField
End Sub

Private Sub lblmyweb_Click()
    ShellExecute hwnd, "open", "http://danlihua.yeah.net", 0, 0, 0
End Sub

Private Sub lblNM_Click()
  lblSerch.Caption = "����" & lblNM.Caption
   DBCombo1.ListField = TxtNM.DataField
End Sub

Private Sub lblNumber_Change()
  Data1.Caption = "���к�" & lblNumber.Caption
End Sub

Private Sub lblPC_Click()
   lblSerch.Caption = "����" & lblPC.Caption
   DBCombo1.ListField = TxtPC.DataField
End Sub

Private Sub lblRegN_Click()
  lblSerch.Caption = "����" & lblRegN.Caption
   DBCombo1.ListField = lblRegNdisplay.DataField
End Sub

Private Sub lblTF_Click()
  lblSerch.Caption = "����" & lblTF.Caption
   DBCombo1.ListField = TxtTF.DataField
End Sub

Private Sub lblUserN_Click()
  lblSerch.Caption = "����" & lblUserN.Caption
   DBCombo1.ListField = txtUserN.DataField
End Sub

Private Sub lblWeb_Click()
lblSerch.Caption = "����" & lblWeb.Caption
   DBCombo1.ListField = txtWeb.DataField
End Sub

Private Sub lblXLH_Click()
 lblSerch.Caption = "�������кţ�"
   DBCombo1.ListField = lblNumber.DataField
End Sub

Private Sub Print_Click()
Printer.EndDoc
End Sub

Private Sub Reginfo_Click()
Frame1.Visible = True
End Sub

Private Sub RegKey1_Finish()
lblRegNdisplay.Caption = RegKey1.RegCode
txtUserN.Text = RegKey1.UserCode
lblDate.Caption = Date
lblAppName.Caption = InputAppname
End Sub
      
Private Sub RegKey2_Finish()

End Sub

Private Sub txtAppN_Change()
Dim msgOK
If RegKey1.User = False Then
txtAppN.Text = CStr(RegKey1.AppCode)
Else
  If Len(txtAppN.Text) = 9 Then
     If txtAppN.Text = "" Or txtAppN.Text = "         " Then txtAppN.Text = "0"
  On Error GoTo errsub
  RegKey1.AppCode = CLng(txtAppN.Text)
  End If
End If
GoTo ed
errsub:
  msgOK = MsgBox("����������")
ed:
End Sub

Private Sub TxtEW_DblClick()
 Dim Email As String
   Email = "mailto:" & TxtEW.Text
   ShellExecute hwnd, "open", Email, 0, 0, 0
End Sub

Private Sub txtName_Change()
RegKey1.AppName = txtName.Text
End Sub

Private Sub txtUserN_Change()
If Len(txtUserN.Text) = 23 Then
RegKey1.UserCode = txtUserN.Text
End If
End Sub

Private Sub txtWeb_DblClick()
   Dim httpwww As String
   httpwww = "http://" & txtWeb.Text
   ShellExecute hwnd, "open", httpwww, 0, 0, 0
End Sub
