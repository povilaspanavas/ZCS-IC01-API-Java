VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "�������������޸�"
   ClientHeight    =   7035
   ClientLeft      =   6135
   ClientTop       =   4005
   ClientWidth     =   11190
   LinkTopic       =   "Form3"
   ScaleHeight     =   7035
   ScaleWidth      =   11190
   Begin VB.Frame Frame2 
      Caption         =   "������"
      Height          =   3615
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   10935
      Begin VB.CommandButton Command6 
         Caption         =   "�ر�"
         Height          =   375
         Left            =   5040
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ִ���޸�ѡ����һ��"
         Height          =   375
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "��Ϣ��"
         Height          =   1215
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   10695
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   77
            Top             =   240
            Width           =   10455
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1200
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "��������"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ѡ����"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   10815
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   15
         Left            =   8040
         TabIndex        =   76
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   14
         Left            =   8040
         TabIndex        =   75
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   13
         Left            =   8040
         TabIndex        =   74
         Top             =   3120
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   12
         Left            =   8040
         TabIndex        =   73
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   11
         Left            =   8040
         TabIndex        =   72
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   10
         Left            =   8040
         TabIndex        =   71
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   9
         Left            =   8040
         TabIndex        =   70
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   8
         Left            =   8040
         TabIndex        =   69
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   7
         Left            =   2640
         TabIndex        =   68
         Top             =   4080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   6
         Left            =   2640
         TabIndex        =   67
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   5
         Left            =   2640
         TabIndex        =   66
         Top             =   3120
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   4
         Left            =   2640
         TabIndex        =   65
         Top             =   2640
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   64
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   63
         Top             =   1680
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   62
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   31
         Left            =   8760
         TabIndex        =   61
         Text            =   "FF FF FF FF FF FF"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   30
         Left            =   6240
         TabIndex        =   60
         Text            =   "FF FF FF FF FF FF"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   29
         Left            =   8760
         TabIndex        =   59
         Text            =   "FF FF FF FF FF FF"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   28
         Left            =   6240
         TabIndex        =   58
         Text            =   "FF FF FF FF FF FF"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   27
         Left            =   8760
         TabIndex        =   57
         Text            =   "FF FF FF FF FF FF"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   26
         Left            =   6240
         TabIndex        =   56
         Text            =   "FF FF FF FF FF FF"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   25
         Left            =   8760
         TabIndex        =   55
         Text            =   "FF FF FF FF FF FF"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   24
         Left            =   6240
         TabIndex        =   54
         Text            =   "FF FF FF FF FF FF"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   23
         Left            =   8760
         TabIndex        =   53
         Text            =   "FF FF FF FF FF FF"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   22
         Left            =   6240
         TabIndex        =   52
         Text            =   "FF FF FF FF FF FF"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   21
         Left            =   8760
         TabIndex        =   51
         Text            =   "FF FF FF FF FF FF"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   20
         Left            =   6240
         TabIndex        =   50
         Text            =   "FF FF FF FF FF FF"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   19
         Left            =   8760
         TabIndex        =   49
         Text            =   "FF FF FF FF FF FF"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   18
         Left            =   6240
         TabIndex        =   48
         Text            =   "FF FF FF FF FF FF"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   17
         Left            =   8760
         TabIndex        =   47
         Text            =   "FF FF FF FF FF FF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   16
         Left            =   6240
         TabIndex        =   46
         Text            =   "FF FF FF FF FF FF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   15
         Left            =   3360
         TabIndex        =   45
         Text            =   "FF FF FF FF FF FF"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   14
         Left            =   840
         TabIndex        =   44
         Text            =   "FF FF FF FF FF FF"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   13
         Left            =   3360
         TabIndex        =   43
         Text            =   "FF FF FF FF FF FF"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   12
         Left            =   840
         TabIndex        =   42
         Text            =   "FF FF FF FF FF FF"
         Top             =   3600
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   11
         Left            =   3360
         TabIndex        =   41
         Text            =   "FF FF FF FF FF FF"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   10
         Left            =   840
         TabIndex        =   40
         Text            =   "FF FF FF FF FF FF"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   9
         Left            =   3360
         TabIndex        =   39
         Text            =   "FF FF FF FF FF FF"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   8
         Left            =   840
         TabIndex        =   38
         Text            =   "FF FF FF FF FF FF"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   7
         Left            =   3360
         TabIndex        =   37
         Text            =   "FF FF FF FF FF FF"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   6
         Left            =   840
         TabIndex        =   36
         Text            =   "FF FF FF FF FF FF"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   3360
         TabIndex        =   35
         Text            =   "FF FF FF FF FF FF"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   840
         TabIndex        =   34
         Text            =   "FF FF FF FF FF FF"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   33
         Text            =   "FF FF FF FF FF FF"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   840
         TabIndex        =   32
         Text            =   "FF FF FF FF FF FF"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   31
         Text            =   "FF FF FF FF FF FF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ȫ��ȡ��"
         Height          =   375
         Left            =   7680
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ȫ��ѡ��"
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�޸�"
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Text            =   "FF FF FF FF FF FF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   "������"
         Height          =   255
         Left            =   9120
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "������"
         Height          =   255
         Left            =   6720
         TabIndex        =   22
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label18 
         Caption         =   "��15��"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "��14��"
         Height          =   255
         Left            =   5640
         TabIndex        =   20
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "��13��"
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "��12��"
         Height          =   255
         Left            =   5640
         TabIndex        =   18
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "��11��"
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "��10��"
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "��09��"
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "��08��"
         Height          =   255
         Left            =   5640
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "������"
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "������"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "��07��"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "��06��"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "��05��"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "��04��"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "��03��"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "��02��"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "��01��"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "��00��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�����ļ�����
Private Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function piccchangesinglekey Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal piccoldkey As Long, ByVal piccnewkey As Long) As Byte
Private Sub Check1_Click(Index As Integer)
If (Check1(Index).Value = 1) Then
     Text1(2 * Index).Enabled = True
     Text1(2 * Index + 1).Enabled = True
Else
    Text1(2 * Index).Enabled = False
    Text1(2 * Index + 1).Enabled = False
End If
End Sub

Private Sub Combo2_Click()
File1.FileName = Combo2.Text + ":\*.ini"
End Sub

Private Sub Command1_Click()
For i = 0 To 15
Check1(i).Value = 1
Next i

For j = 0 To 31
Text1(j).Enabled = True
Next j
End Sub

Private Sub Command2_Click()
For i = 0 To 15
Check1(i).Value = 0
Next i

For j = 0 To 31
Text1(j).Enabled = False
Next j

End Sub

Private Sub Command3_Click()
Dim update As String  '���������޸�����ͳ�ƿ����
Dim checkCount As Long
For i = 0 To 15
    If (Check1(i).Value <> 0) Then
    checkCount = checkCount + 1
    End If
Next i
If (checkCount = 0) Then
Text4.Text = "û��ѡ���κ���Ҫ�޸������������ѡ��"
Else
For i = 0 To 15
    If (Check1(i).Value = 1) Then
        '�޸�����
 Dim status As Byte '��ŷ���ֵ
Dim myareano As Byte '����
Dim authmode As Byte '�������ͣ���A�����B����
Dim myctrlword As Byte '������
Dim mypiccserial(0 To 3) As Byte '�����к�
Dim mypiccoldkey(0 To 5) As Byte '������
Dim mypiccnewkey(0 To 5) As Byte '������
myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + EXTERNKEY
 If (Combo1.ListIndex = 0) Then
  authmode = 1
  Else
  authmode = 0
 End If
 
    '
     'ָ��������
  Dim str1() As String
  Dim strLength1 As Integer
  Dim flag1 As Boolean
  str1 = Split(Text1(i).Text, " ")
 strLength1 = UBound(str1) - LBound(str1) + 1
 If (strLength1 = 6) Then
 For m = 0 To (strLength1 - 1)
     If (Len(str1(m)) <> 2) Then
     Text4.Text = "��ʽ����!"
     flag1 = True
     Exit For
    End If
 Next m
 Else
  flag1 = True
 Text4.Text = "�����������ʽ����!"
 End If

If (flag1 = False) Then
For n = 0 To strLength1 - 1
 mypiccoldkey(n) = "&H" & str1(n)
Next n
End If
'ָ��������
  Dim str2() As String
  Dim strLength2 As Integer
  Dim flag2 As Boolean
  str2 = Split(Text1(i + 1).Text, " ")
 strLength2 = UBound(str2) - LBound(str2) + 1
 If (strLength2 = 6) Then
 For o = 0 To (strLength2 - 1)
     If (Len(str2(o)) <> 2) Then
     Text4.Text = "��ʽ����!"
     flag2 = True
     Exit For
    End If
 Next o
 Else
  flag2 = True
 Text4.Text = "�����������ʽ����!"
 End If

If (flag2 = False) Then
For p = 0 To strLength2 - 1
 mypiccnewkey(p) = "&H" & str2(p)
Next p
End If
'
    status = piccchangesinglekey(myctrlword, VarPtr(mypiccserial(0)), i, authmode, VarPtr(mypiccoldkey(0)), VarPtr(mypiccnewkey(0)))
    If (status = 0) Then
    update = update + "[" + CStr(i) + "] "
    Else
    update = update + CStr(i) + " "
    End If
'�޸�����
    End If
Next i
    Text4.Text = "�����ɹ�!" + vbCrLf + "�����к�:" + CStr(Hex(mypiccserial(0))) + "-" + CStr(Hex(mypiccserial(1))) + "-" + CStr(Hex(mypiccserial(2))) + "-" + CStr(Hex(mypiccserial(3))) + vbCrLf + update + " ��[]���Ϊ�޸ĳɹ�������"
    End If

End Sub

Private Sub Command4_Click()
Dim mark As Long
Dim result As String
Dim pass() As String
Dim passLength As Long
Dim devil As String
Dim haha As Boolean
Dim counts As Integer
result = String(287, 0)
If (File1.FileName = Null Or File1.FileName = "") Then
MsgBox ("��ѡ��Ҫ������ļ�!")
Else
mark = GetPrivateProfileStringA("password", "value", "", result, 288, Combo2.Text + ":\" + File1.FileName)
pass = Split(result, " ")
passLength = UBound(pass) - LBound(pass) + 1
devil = "[0-9A-Fa-f]{2}"
For i = 0 To passLength - 1
haha = bTest(pass(i), devil)
    If (haha = False) Then
    Text4.Text = "�����ļ��к��в�����ʮ������(0~9,a~f,A~F)���ַ��������޸ĺ��ٵ���!"
    counts = counts + 1
    Exit For
    End If
    If (Len(pass(i)) <> 2) Then
    Text4.Text = "�����ʽ����Ӧ����������xx xx xx...�����޸ĺ��ٵ���!"
    counts = counts + 1
    Exit For
    End If
Next i
If (passLength < 96) Then
Text4.Text = "�����ܳ��Ȳ��������޸ĺ��ٵ���!"
counts = counts + 1
End If
If (counts = 0) Then
'pass = Split(result, " ")
    For i = 0 To 30
        If (i Mod 2 = 0) Then
        For j = 0 To 5
        Dim flag As String
       flag = flag + pass(j + i * 3) + " "
        Next j
          Text1(i).Text = RTrim(flag)
          Text1(i).Enabled = True
          flag = ""
        End If
    Next i
Text4.Text = "����ɹ�!"
End If
End If
End Sub

Private Sub Command6_Click()
Unload Form3
Form1.Visible = True
End Sub

Private Sub Form_Load()
Combo1.AddItem ("A")
Combo1.AddItem ("B")
Combo1.ListIndex = 0
For i = 0 To 31
Text1(i).Enabled = False
Text1(i).MaxLength = 17
Next i

End Sub


Function bTest(ByVal s As String, ByVal p As String) As Boolean
    Dim re As RegExp
    Set re = New RegExp
    re.IgnoreCase = False  '�����Ƿ�ƥ���Сд
    re.Pattern = p
    bTest = re.Test(s)
End Function


