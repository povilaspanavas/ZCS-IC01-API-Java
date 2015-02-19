VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "数据读写"
   ClientHeight    =   6630
   ClientLeft      =   6840
   ClientTop       =   4005
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   9015
   Begin VB.Timer Timer1 
      Left            =   6600
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operating area"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.Frame Frame3 
         Caption         =   "Information"
         Height          =   1455
         Left            =   120
         TabIndex        =   24
         Top             =   4560
         Width           =   8295
         Begin VB.CommandButton Command4 
            Caption         =   "exit"
            Height          =   495
            Left            =   6480
            TabIndex        =   26
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox Text5 
            Height          =   855
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   6000
         TabIndex        =   21
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Card data"
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   8175
         Begin VB.ComboBox Combo5 
            Height          =   300
            Left            =   5760
            TabIndex        =   28
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Height          =   615
            Left            =   6720
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Height          =   615
            Left            =   5760
            TabIndex        =   23
            Top             =   1800
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Height          =   615
            Left            =   5760
            TabIndex        =   22
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   600
            TabIndex        =   20
            Text            =   "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
            Top             =   2040
            Width           =   4575
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   600
            TabIndex        =   18
            Text            =   "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
            Top             =   1560
            Width           =   4575
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   600
            TabIndex        =   16
            Text            =   "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Text            =   "FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF FF"
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label11 
            Caption         =   "/S"
            Height          =   255
            Left            =   6480
            TabIndex        =   29
            Top             =   600
            Width           =   255
         End
         Begin VB.Image Image4 
            Height          =   300
            Left            =   5160
            Picture         =   "Form2.frx":0000
            Top             =   2040
            Width           =   435
         End
         Begin VB.Image Image3 
            Height          =   300
            Left            =   5160
            Picture         =   "Form2.frx":0392
            Top             =   1560
            Width           =   435
         End
         Begin VB.Image Image2 
            Height          =   300
            Left            =   5160
            Picture         =   "Form2.frx":0724
            Top             =   1080
            Width           =   435
         End
         Begin VB.Image Image1 
            Height          =   300
            Left            =   5160
            Picture         =   "Form2.frx":0AB6
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label10 
            Caption         =   "block3"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "block2"
            Height          =   255
            Left            =   0
            TabIndex        =   17
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "block1"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "block0"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   " 0  1  2  3  4  5  6  7  8  9  A  B  C  D  E  F"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   2280
         Picture         =   "Form2.frx":0E48
         ScaleHeight     =   315
         ScaleWidth      =   1515
         TabIndex        =   9
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   3360
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   840
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   5160
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   1800
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Top             =   3720
         Width           =   15
      End
      Begin VB.Label Label4 
         Caption         =   "New pass"
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Old pass"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "password type:"
         Height          =   255
         Left            =   3720
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "current area:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private counts As Long '计数变量
Dim countFlag As Long '连续测试按钮标识
Private Const BLOCK0_EN = &H1
Private Const BLOCK1_EN = &H2
Private Const BLOCK2_EN = &H4
Private Const NEEDSERIAL = &H8
Private Const EXTERNKEY = &H10
'一次性读取单区0~2块函数
Private Declare Function piccreadex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal key As Byte, ByVal picckey As Long, ByVal piccdata As Long) As Byte
'写单区函数
Private Declare Function piccwriteex Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long, ByVal piccdata0_2 As Long) As Byte
'读取设备编号函数声明
Private Declare Function pcdgetdevicenumber Lib "OUR_MIFARE.dll" (ByVal devicenumber As Long) As Byte
'寻卡序列号
Private Declare Function piccrequest Lib "OUR_MIFARE.dll" (ByVal serial As Long) As Byte
'密码认证方式1
Private Declare Function piccauthkey1 Lib "OUR_MIFARE.dll" (ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal picckey As Long) As Byte
'单块函数读取
Private Declare Function piccread Lib "OUR_MIFARE.dll" (ByVal block As Byte, ByVal piccdata As Long) As Byte
'修改单区密码声明
Private Declare Function piccchangesinglekey Lib "OUR_MIFARE.dll" (ByVal ctrlword As Byte, ByVal serial As Long, ByVal area As Byte, ByVal keyA1B0 As Byte, ByVal piccoldkey As Long, ByVal piccnewkey As Long) As Byte

Private Sub Command1_Click()
Dim status As Byte '存放返回值
Dim myareano As Byte '区号
Dim authmode As Byte '密码类型，用A密码或B密码
Dim myctrlword As Byte '控制字
Dim mypiccserial(0 To 3) As Byte '卡序列号
Dim mypiccoldkey(0 To 5) As Byte '旧密码
Dim mypiccnewkey(0 To 5) As Byte '新密码
myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + EXTERNKEY
 'myareano = CByte(Combo1.ListIndex) * 4
 myareano = CByte(Combo1.ListIndex)
 If (Combo2.ListIndex = 0) Then
  authmode = 1
  Else
  authmode = 0
 End If
 '指定旧密码
  Dim str1() As String
  Dim strLength1 As Integer
  Dim flag1 As Boolean
  str1 = Split(Combo3.Text, " ")
 strLength1 = UBound(str1) - LBound(str1) + 1
 If (strLength1 = 6) Then
 For i = 0 To (strLength1 - 1)
     If (Len(str1(i)) <> 2) Then
     Text5.Text = "Format error!"
     flag1 = True
     Exit For
    End If
 Next i
 Else
  flag1 = True
 Text5.Text = "Old Password command format error!"
 End If

If (flag1 = False) Then
For i = 0 To strLength1 - 1
 mypiccoldkey(i) = "&H" & str1(i)
Next
End If
'指定新密码
  Dim str2() As String
  Dim strLength2 As Integer
  Dim flag2 As Boolean
  str2 = Split(Combo4.Text, " ")
 strLength2 = UBound(str2) - LBound(str2) + 1
 If (strLength2 = 6) Then
 For i = 0 To (strLength2 - 1)
     If (Len(str2(i)) <> 2) Then
     Text5.Text = "Format error!"
     flag2 = True
     Exit For
    End If
 Next i
 Else
  flag2 = True
 Text5.Text = "New Password command format error!"
 End If

If (flag2 = False) Then
For i = 0 To strLength2 - 1
 mypiccnewkey(i) = "&H" & str2(i)
Next
End If

If (flag1 = False And flag2 = False) Then
status = piccchangesinglekey(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypiccoldkey(0)), VarPtr(mypiccnewkey(0)))
'处理返回函数
Select Case status
    Case 0:
        Text5.Text = "Area " + CStr(Combo1.ListIndex) + ",Password changed successfully!"
        Combo3.AddItem (Combo4.Text)
        Combo4.AddItem (Combo4.Text)
    Case 8:
         Text5.Text = "Keep the card on the sensor area"
    Case 9:
         Text5.Text = "Multiple cards in the sensing area!"
    Case 10:
         Text5.Text = "The card may have been dormant!"
    Case 11:
         Text5.Text = "Password loading fails!"
    Case 12:
         Text5.Text = "Please enter the correct password and retry!"
    Case 21 '没有动态库
         Text5.Text = "找不到动态库ICUSB.DLL请将ICUSB.DLL拷贝到VB安装后的目录VB98下"
    Case Else
         Text5.Text = "Exception"

End Select
Else
Text5.Text = "Operation fails!"
End If
End Sub

Private Sub Command2_Click()
  Dim status As Byte '存放返回值
  Dim myareano As Byte '区号
  Dim authmode As Byte '密码类型，用A密码或B密码
  Dim myctrlword As Byte '控制字
  Dim mypicckey(0 To 5) As Byte '密码
  Dim mypiccserial(0 To 3) As Byte '卡序列号
  Dim mypiccdata(0 To 63) As Byte '卡数据缓冲
  Dim cardNumber As String '结果-序列号
  Dim block1 As String
  Dim block2 As String
  Dim block3 As String
  Dim imageFlag As Boolean
  Dim str() As String
  Dim strLength As Integer
  Dim flag2 As Boolean
  'myareano = CByte(Combo1.ListIndex) * 4
  myareano = CByte(Combo1.ListIndex)
 If (Combo2.ListIndex = 0) Then
  authmode = 1
  Else
  authmode = 0
 End If
 
 str = Split(Combo3.Text, " ")
 strLength = UBound(str) - LBound(str) + 1
 If (strLength = 6) Then
 For i = 0 To (strLength - 1)
     If (Len(str(i)) <> 2) Then
     Text5.Text = "Format error!"
     flag2 = True
     Exit For
    End If
 Next i
 Else
  flag2 = True
 Text5.Text = "Please enter the six commands!"
 End If

If (flag2 = False) Then
For i = 0 To strLength - 1
 mypicckey(i) = "&H" & str(i)
Next
End If
  'mypicckey(0) = &HFF
  'mypicckey(1) = &HFF
  'mypicckey(2) = &HFF
  'mypicckey(3) = &HFF
  'mypicckey(4) = &HFF
  'mypicckey(5) = &HFF
  myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + NEEDSERIAL + EXTERNKEY
  status = piccrequest(VarPtr(mypiccserial(0)))
'status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
'For i = 0 To 3
'mark = mark + IIf(Len(Hex(mypiccserial(i))) < 2, "0" + Hex(mypiccserial(i)), Hex(mypiccserial(i))) + " "
'Next i
Select Case status
Case 0:
status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
For i = 0 To 3
mark = mark + IIf(Len(Hex(mypiccserial(i))) < 2, "0" + Hex(mypiccserial(i)), Hex(mypiccserial(i))) + " "
Next i

For i = 0 To 47
   If (i < 15) Then
   block1 = block1 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 15) Then
   block1 = block1 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   ElseIf (15 < i And i < 31) Then
   block2 = block2 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 31) Then
   block2 = block2 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   ElseIf (31 < i And i < 47) Then
   block3 = block3 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 47) Then
   block3 = block3 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   End If
Next i
imageFlag = True
Text1.Text = block1
Text2.Text = block2
Text3.Text = block3
'分步读取第3块
Dim piccdata(0 To 15) As Byte
Dim block As Byte
Dim block4 As String
block = myareano * 4 + 3
    status = piccrequest(VarPtr(mypiccserial(0)))
    status = piccauthkey1(VarPtr(mypiccserial(0)), block, authmode, VarPtr(mypicckey(0)))
    status = piccread(block, VarPtr(piccdata(0)))
For i = 0 To 15
block4 = block4 + IIf(Len(Hex(piccdata(i))) < 2, "0" + Hex(piccdata(i)), Hex(piccdata(i))) + " "
Next i
Text4.Text = block4
'分步读取第3块
counts = counts + 1
Text5.Text = "Card serial number：" + mark + ",Area " + CStr(Combo1.ListIndex) + "," + "block 0~3 Read successfully!...< times " + CStr(counts) + ">"
Case 1:
    imageFlag = False
    Text5.Text = "0~2块都没读出来，可能刷卡太块!"
Case 2:
    imageFlag = False
    Text5.Text = "第0块已被读出，但1~2块读取失败!"
Case 3:
    imageFlag = False
    Text5.Text = "第0、1块已被读出，但2块读取失败!"
Case 8:
    imageFlag = False
    Text5.Text = "Keep the card on the sensor area!"
Case 9:
    imageFlag = False
    Text5.Text = "Multiple cards in the sensing area!"
Case 10:
    imageFlag = False
    Text5.Text = "该卡可能已被休眠，无法选中!"
Case 11:
    imageFlag = False
    Text5.Text = "密码装载失败!"
Case 12:
    imageFlag = False
    Text5.Text = "Password authentication failed,Please enter the correct password and retry!"
Case 21:
    imageFlag = False
    Text5.Text = "找不到动态库ICUSB.DLL请将ICUSB.DLL拷贝到VB安装后的目录VB98下"
Case Else
    imageFlag = False
    Text5.Text = "exception"
End Select
If (imageFlag) Then
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Else
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End If
End Sub
Private Sub Command3_Click()
Dim status As Byte '存放返回值
Dim myareano As Byte '区号
Dim authmode As Byte '密码类型，用A密码或B密码
Dim myctrlword As Byte '控制字
Dim mypicckey(0 To 5) As Byte '密码
Dim mypiccserial(0 To 3) As Byte '卡序列号
Dim mypiccdata(0 To 47) As Byte '卡数据缓冲
myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + EXTERNKEY
  'myareano = CByte(Combo1.ListIndex) * 4
  myareano = CByte(Combo1.ListIndex)
 If (Combo2.ListIndex = 0) Then
  authmode = 1
  Else
  authmode = 0
 End If
 
  Dim str() As String
  Dim strLength As Integer
  Dim flag2 As Boolean
   str = Split(Combo3.Text, " ")
 strLength = UBound(str) - LBound(str) + 1
 If (strLength = 6) Then
 For i = 0 To (strLength - 1)
     If (Len(str(i)) <> 2) Then
     Text5.Text = "Format error!"
     flag2 = True
     Exit For
    End If
 Next i
 Else
  flag2 = True
 Text5.Text = "Please enter the six commands!"
 End If

If (flag2 = False) Then
For i = 0 To strLength - 1
 mypicckey(i) = "&H" & str(i)
Next
End If
'指定卡数据
Dim block1() As String
Dim block2() As String
Dim block3() As String
block1 = Split(Text1.Text, " ")
block2 = Split(Text2.Text, " ")
block3 = Split(Text3.Text, " ")
Dim lengFlag As Long
Dim block1Length As Integer
Dim block2Length As Integer
Dim block3Length As Integer

'写入数据格式判断
 block1Length = UBound(block1) - LBound(block1) + 1
 If (block1Length = 16) Then
 For i = 0 To (block1Length - 1)
     If (Len(block1(i)) <> 2) Then
     Text5.Text = "Data format error, please check and re-enter!"
     lengFlag = lengFlag + 1
     Exit For
    End If
 Next i
 Else
 Text5.Text = "The length of the data in error, please re-enter!"
  lengFlag = lengFlag + 1
 End If
 
  block2Length = UBound(block2) - LBound(block2) + 1
 If (block2Length = 16) Then
 For i = 0 To (block2Length - 1)
     If (Len(block2(i)) <> 2) Then
     Text5.Text = "Data format error, please check and re-enter!"
     lengFlag = lengFlag + 1
     Exit For
    End If
 Next i
 Else
 Text5.Text = "The length of the data in error, please re-enter!"
  lengFlag = lengFlag + 1
 End If

 block3Length = UBound(block3) - LBound(block3) + 1
 If (block3Length = 16) Then
 For i = 0 To (block3Length - 1)
     If (Len(block3(i)) <> 2) Then
     Text5.Text = "Data format error, please check and re-enter!"
     lengFlag = lengFlag + 1
     Exit For
    End If
 Next i
 Else
 Text5.Text = "The length of the data in error, please re-enter!"
  lengFlag = lengFlag + 1
 End If
'

If (lengFlag = 0) Then
For i = 0 To 47
    If (i < 16) Then
    mypiccdata(i) = "&H" + block1(i)
    ElseIf (15 < i And i < 32) Then
    mypiccdata(i) = "&H" + block2(i - 16)
    ElseIf (31 < i And i < 48) Then
    mypiccdata(i) = "&H" + block3(i - 32)
    End If
Next i
status = piccwriteex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
Select Case status
    Case 0:
    Dim mark As String
    For i = 0 To 3
    mark = mark + IIf(Len(Hex(mypiccserial(i))) < 2, "0" + Hex(mypiccserial(i)), Hex(mypiccserial(i))) + " "
    Next i
        Text5.Text = "Card serial number：" + mark + ",Area " + CStr(Combo1.ListIndex) + "," + "block 0~3 Write card success!"
    Case 8:
        Text5.Text = "Keep the card on the sensor area"
    Case 21 '没有动态库
        Text5.Text = "Can not find dynamic library ICUSB.DLL, your ICUSB.DLL copy to the the VB installation directory VB98"
    Case Else
        Text5.Text = "Exception"
End Select
End If
End Sub
Private Sub Command5_Click()
Dim times As Long
countFlag = countFlag + 1
times = (CLng(Combo5.Text)) * 1000
Timer1.Interval = times
    If (countFlag Mod 2 = 0) Then
        Timer1.Enabled = False
        Command5.Caption = "Continuous read area " + CStr(Combo1.ListIndex) + "(R)"
    Else
        Command5.Caption = "Stop"
        counts = 0
        Timer1.Enabled = True
    End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("area 00")
Combo1.AddItem ("area 01")
Combo1.AddItem ("area 02")
Combo1.AddItem ("area 03")
Combo1.AddItem ("area 04")
Combo1.AddItem ("area 05")
Combo1.AddItem ("area 06")
Combo1.AddItem ("area 07")
Combo1.AddItem ("area 08")
Combo1.AddItem ("area 09")
Combo1.AddItem ("area 10")
Combo1.AddItem ("area 11")
Combo1.AddItem ("area 12")
Combo1.AddItem ("area 13")
Combo1.AddItem ("area 14")
Combo1.AddItem ("area 15")
Combo1.ListIndex = 0
Combo2.AddItem ("A")
Combo2.AddItem ("B")
Combo2.ListIndex = 0
Combo3.AddItem ("FF FF FF FF FF FF")
Combo3.ListIndex = 0
Combo4.AddItem ("FF FF FF FF FF FF")
Combo4.ListIndex = 0
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False

Timer1.Enabled = False
Combo5.AddItem (1)
Combo5.AddItem (1.5)
Combo5.AddItem (2)
Combo5.AddItem (2.5)
Combo5.AddItem (3)
Combo5.AddItem (3.5)
Combo5.AddItem (4)
Combo5.AddItem (4.5)
Combo5.AddItem (5)
Combo5.ListIndex = 0
Command5.Caption = "Continuous read area " + CStr(Combo1.ListIndex) + "(R)"
End Sub

Private Sub Combo1_Click()
Command1.Caption = "Modify area " + CStr(Combo1.ListIndex) + " password"
Command2.Caption = "Read area" + CStr(Combo1.ListIndex) + "(R)"
Command3.Caption = "Write area" + CStr(Combo1.ListIndex) + "(W)"
If (Timer1.Enabled = False) Then
Command5.Caption = "Continuous read area " + CStr(Combo1.ListIndex) + "(R)"
End If
End Sub

Private Sub Command4_Click()
Unload Form2
Form1.Visible = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 32, Asc("0") To Asc("9")
       Case Asc("a") To Asc("f")
       Case Asc("A") To Asc("F")
       Case Else
        KeyAscii = 0
      Text5.Text = "Your input errors!"
End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 32, Asc("0") To Asc("9")
       Case Asc("a") To Asc("f")
       Case Asc("A") To Asc("F")
       Case Else
        KeyAscii = 0
      Text5.Text = "Your input errors!"
End Select
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 32, Asc("0") To Asc("9")
       Case Asc("a") To Asc("f")
       Case Asc("A") To Asc("F")
       Case Else
        KeyAscii = 0
      Text5.Text = "Your input errors!"
End Select
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
       Case 8, 32, Asc("0") To Asc("9")
       Case Asc("a") To Asc("f")
       Case Asc("A") To Asc("F")
       Case Else
        KeyAscii = 0
      Text5.Text = "Your input errors!"
End Select
End Sub
Private Sub Timer1_Timer()
  Dim status As Byte '存放返回值
  Dim myareano As Byte '区号
  Dim authmode As Byte '密码类型，用A密码或B密码
  Dim myctrlword As Byte '控制字
  Dim mypicckey(0 To 5) As Byte '密码
  Dim mypiccserial(0 To 3) As Byte '卡序列号
  Dim mypiccdata(0 To 63) As Byte '卡数据缓冲
  Dim cardNumber As String '结果-序列号
  Dim block1 As String
  Dim block2 As String
  Dim block3 As String
  Dim imageFlag As Boolean
  Dim str() As String
  Dim strLength As Integer
  Dim flag2 As Boolean
  myareano = CByte(Combo1.ListIndex)
 If (Combo2.ListIndex = 0) Then
  authmode = 1
  Else
  authmode = 0
 End If
 
 str = Split(Combo3.Text, " ")
 strLength = UBound(str) - LBound(str) + 1
 If (strLength = 6) Then
 For i = 0 To (strLength - 1)
     If (Len(str(i)) <> 2) Then
     Text5.Text = "Format error!"
     flag2 = True
     Exit For
    End If
 Next i
 Else
  flag2 = True
 Text5.Text = "Please enter the six commands!"
 End If

If (flag2 = False) Then
For i = 0 To strLength - 1
 mypicckey(i) = "&H" & str(i)
Next
End If
  myctrlword = BLOCK0_EN + BLOCK1_EN + BLOCK2_EN + NEEDSERIAL + EXTERNKEY
  status = piccrequest(VarPtr(mypiccserial(0)))
'status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
'For i = 0 To 3
'mark = mark + IIf(Len(Hex(mypiccserial(i))) < 2, "0" + Hex(mypiccserial(i)), Hex(mypiccserial(i))) + " "
'Next i
Select Case status
Case 0:
'调整读卡数据位置
status = piccreadex(myctrlword, VarPtr(mypiccserial(0)), myareano, authmode, VarPtr(mypicckey(0)), VarPtr(mypiccdata(0)))
For i = 0 To 3
mark = mark + IIf(Len(Hex(mypiccserial(i))) < 2, "0" + Hex(mypiccserial(i)), Hex(mypiccserial(i))) + " "
Next i

For i = 0 To 47
   If (i < 15) Then
   block1 = block1 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 15) Then
   block1 = block1 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   ElseIf (15 < i And i < 31) Then
   block2 = block2 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 31) Then
   block2 = block2 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   ElseIf (31 < i And i < 47) Then
   block3 = block3 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i))) + " "
   ElseIf (i = 47) Then
   block3 = block3 + IIf(Len(Hex(mypiccdata(i))) < 2, "0" + Hex(mypiccdata(i)), Hex(mypiccdata(i)))
   End If
Next i
imageFlag = True
Text1.Text = block1
Text2.Text = block2
Text3.Text = block3
'分步读取第3块
Dim piccdata(0 To 15) As Byte
Dim block As Byte
Dim block4 As String
block = myareano * 4 + 3
    status = piccrequest(VarPtr(mypiccserial(0)))
    status = piccauthkey1(VarPtr(mypiccserial(0)), block, authmode, VarPtr(mypicckey(0)))
    status = piccread(block, VarPtr(piccdata(0)))
For i = 0 To 15
block4 = block4 + IIf(Len(Hex(piccdata(i))) < 2, "0" + Hex(piccdata(i)), Hex(piccdata(i))) + " "
Next i
Text4.Text = block4
'分步读取第3块
counts = counts + 1
Text5.Text = "Card serial：" + mark + ",area " + CStr(Combo1.ListIndex) + ",block" + "0~3 read successfully!...<times " + CStr(counts) + ">"
Case 1:
    imageFlag = False
    Text5.Text = "0~2块都没读出来，可能刷卡太块!"
Case 2:
    imageFlag = False
    Text5.Text = "第0块已被读出，但1~2块读取失败!"
Case 3:
    imageFlag = False
    Text5.Text = "第0、1块已被读出，但2块读取失败!"
Case 8:
    imageFlag = False
    Text5.Text = "Keep the card on the sensor area!"
Case 9:
    imageFlag = False
    Text5.Text = "Keep the card on the sensor area!"
Case 10:
    imageFlag = False
    Text5.Text = "该卡可能已被休眠，无法选中!"
Case 11:
    imageFlag = False
    Text5.Text = "Password loading fails!"
Case 12:
    imageFlag = False
    Text5.Text = "Password authentication failed,Please enter the correct password and retry!"
Case 21:
    imageFlag = False
    Text5.Text = "Can not find dynamic library ICUSB.DLL, your ICUSB.DLL copy to the the VB installation directory VB98"
Case Else
    imageFlag = False
    Text5.Text = "Exception"
End Select
If (imageFlag) Then
Image1.Visible = True
Image2.Visible = True
Image3.Visible = True
Image4.Visible = True
Else
Image1.Visible = False
Image2.Visible = False
Image3.Visible = False
Image4.Visible = False
End If
End Sub
