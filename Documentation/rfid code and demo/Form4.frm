VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "CPU��Ӧ��"
   ClientHeight    =   5400
   ClientLeft      =   7020
   ClientTop       =   5085
   ClientWidth     =   9285
   LinkTopic       =   "Form4"
   ScaleHeight     =   5400
   ScaleWidth      =   9285
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "��Ϣ��"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin VB.TextBox Text1 
         Height          =   2415
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   8535
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ͨ���ź�������
Private Declare Sub IccClose Lib "Lib980.dll" (ByVal slot As Byte)
Private Declare Function IccInit Lib "Lib980.dll" (ByVal slot As Byte, ByVal ATR As Long) As Byte
Private Sub Command1_Click()
Dim slot As Byte
slot = 5
Form1.Visible = True
IccClose (slot)
Unload Form4
End Sub
Private Sub Form_Load()
    Dim slot As Byte
    Dim mark As Byte
    Dim ATR(0 To 32) As Byte
    slot = 5
    mark = IccInit(slot, VarPtr(ATR(0)))
    If (mark = 0) Then
        Text1.Text = "��ʼ���ɹ�!"
    Else
        Text1.Text = "��ʼ��ʧ��!"
    End If
End Sub
