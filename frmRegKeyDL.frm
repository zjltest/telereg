VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmDL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ����Ϣ��"
   ClientHeight    =   2295
   ClientLeft      =   3495
   ClientTop       =   3495
   ClientWidth     =   5130
   Icon            =   "frmRegKeyDL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5130
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "�����"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "123456789"
      Top             =   1640
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "������"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   270
      Left            =   3480
      TabIndex        =   6
      Text            =   "DEMO"
      Top             =   1160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "���к�"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   270
      Left            =   3480
      TabIndex        =   5
      Top             =   680
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "����������Ϣ��"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDBCtls.DBCombo DBCombo 
      Bindings        =   "frmRegKeyDL.frx":030A
      Height          =   330
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   327680
      ListField       =   "������"
      Text            =   ""
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "����ţ�"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "��������"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "���кţ�"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ѡ��Ҫ��¼���û���Ϣ�⡣"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Unload Me
Data1.Recordset.Bookmark = DBCombo.SelectedItem
DLH = Text1.Text
InputAppname = Text2.Text
InputAppcode = Text3.Text
End Sub

Private Sub DBCombo_KeyUp(KeyCode As Integer, Shift As Integer)
If Len(DBCombo.Text) = 9 And DBCombo.Text = Text3.Text Then
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text3.PasswordChar = ""
Label5.Caption = "����" & Label3.Caption
DBCombo.ListField = Text2.DataField
End If
Data1.Recordset.Bookmark = DBCombo.SelectedItem
DLH = Text1.Text
InputAppname = Text2.Text
InputAppcode = Text3.Text
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\RegKey.mdb"
If IfEdit = True Then
Me.Caption = "�༭������Ϣ��"
Label1.Caption = "�������Ӧ�ó���������Ϣ��ע�Ᵽ�ܣ�"
Label5.Caption = "Ϊ�����������ݣ������뵱ǰ����š�"
Else
Label5.Caption = "����" & Label3.Caption
DBCombo.ListField = Text2.DataField
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Data1.Recordset.Bookmark = DBCombo.SelectedItem
DLH = Text1.Text
InputAppname = Text2.Text
InputAppcode = Text3.Text
Unload Me
End Sub

Private Sub Label2_Click()
Label5.Caption = "����" & Label2.Caption
   DBCombo.ListField = Text1.DataField
End Sub

Private Sub Label3_Click()
Label5.Caption = "����" & Label3.Caption
   DBCombo.ListField = Text2.DataField
End Sub

