VERSION 5.00
Object = "{6E0B417E-61E2-11D2-8F6D-D0E44AC10000}#7.0#0"; "TELEREG.OCX"
Begin VB.Form TeleRegDemo 
   Caption         =   "TeleRegDemo"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5775
      Begin CtrTeleReg.TeleReg TeleReg1 
         Height          =   1455
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   2566
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483630
         ForeColor       =   -2147483630
         ForeColor       =   -2147483630
         ForeColor       =   -2147483630
         ForeColor       =   -2147483630
         MsgInfo         =   -1  'True
         Interval        =   60000
      End
      Begin VB.Label Label1 
         Caption         =   "请注册"
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   4455
      End
   End
End
Attribute VB_Name = "TeleRegDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

       Private Sub TeleReg1_CancelClick()
         Dim Msgok
           Msgok = MsgBox("你还没有注册，你想以后再注册？", _
                  vbQuestion, "以后注册？")
           Frame1.Visible = False
       End Sub

       Private Sub TeleReg1_Userfalse()
             Frame1.Visible = True
       End Sub

       Private Sub TeleReg1_Usertrue()
            Dim Msgok
           Msgok = MsgBox("恭喜你注册成功，欢迎使用正版软件！", _
                   vbMsgok, "注册成功")
           Frame1.Visible = False
       End Sub

       Private Sub TeleReg1_Nofree()
            End
       End Sub

