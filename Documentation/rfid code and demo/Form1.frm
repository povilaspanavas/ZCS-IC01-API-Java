VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "批量密码修改器V1.0"
   ClientHeight    =   2940
   ClientLeft      =   7740
   ClientTop       =   5625
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   7755
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton Command3 
         Caption         =   "exit"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "change password"
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "read and write"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "非接触式感应卡批量密码修改器示例程序V1.0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function pcdgetdevicenumber Lib "OUR_MIFARE.dll" (ByVal devicenumber As Long) As Byte

Private Sub Command1_Click()
Form1.Visible = False
Form2.Show
End Sub

Private Sub Command2_Click()
Form3.Show
Form1.Visible = False
End Sub

Private Sub Command3_Click()
Unload Form1
End Sub
Private Sub Command4_Click()
Form4.Show
Form1.Visible = False
End Sub

Private Sub Form_Load()

Dim mark As Byte
Dim result As String
Dim devicenumber(0 To 3) As Byte
mark = pcdgetdevicenumber(VarPtr(devicenumber(0)))
If (mark = 0) Then
For i = 0 To 3
    If (i = 3) Then
     result = result + Hex(devicenumber(i))
    Else
    result = result + Hex(devicenumber(i)) + "-"
    End If
Next i
Label1.Caption = "设备全球唯一序列编号：" + result
Else
Command1.Enabled = False
Command2.Enabled = False
Label1.Caption = "设备连接错误、请重试..."
End If

End Sub
