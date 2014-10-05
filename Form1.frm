VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12195
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":4D5A
   ScaleHeight     =   7860
   ScaleWidth      =   12195
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox PicData 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   9810
      ScaleHeight     =   2445
      ScaleWidth      =   2325
      TabIndex        =   4
      Top             =   4785
      Width           =   2325
      Begin VB.TextBox Text1 
         Height          =   345
         Index           =   0
         Left            =   1155
         TabIndex        =   7
         Top             =   765
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Index           =   1
         Left            =   1170
         TabIndex        =   6
         Top             =   1365
         Width           =   1100
      End
      Begin VB.TextBox Text1 
         Height          =   345
         Index           =   2
         Left            =   1170
         TabIndex        =   5
         Top             =   1965
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "【数据区】"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   540
         TabIndex        =   11
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "16进制值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   45
         TabIndex        =   10
         Top             =   2010
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RGB 值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   285
         TabIndex        =   9
         Top             =   1410
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "十进制值:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   45
         TabIndex        =   8
         Top             =   810
         Width           =   1080
      End
   End
   Begin VB.PictureBox PicPhoto 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   5475
      Picture         =   "Form1.frx":101780
      ScaleHeight     =   3150
      ScaleWidth      =   4200
      TabIndex        =   3
      Top             =   105
      Width           =   4200
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   120
      Top             =   4830
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   105
      Top             =   4290
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选择图片"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label LabPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   10935
      TabIndex        =   16
      Top             =   4425
      Width           =   120
   End
   Begin VB.Label LabPos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   10935
      TabIndex        =   15
      Top             =   4005
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "窗体坐标:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   9795
      TabIndex        =   14
      Top             =   4425
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "屏幕坐标:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   9795
      TabIndex        =   13
      Top             =   4005
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "【鼠标坐标位置】"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   9855
      TabIndex        =   12
      Top             =   3600
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "【2】您可以在其它地方截取图片复制到剪贴板内,然后再在图片显示区内使用鼠标右键点击即可以将图像黏贴进图像显示区域内"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Index           =   1
      Left            =   9855
      TabIndex        =   2
      Top             =   1665
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "【1】鼠标左键点击屏幕上任意点想要的颜色,RGB值将显示在下面【数据区】三个文本框内"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Index           =   0
      Left            =   9855
      TabIndex        =   1
      Top             =   210
      Width           =   2160
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'摘录自【CBM666 VB编程示例教材 图像篇_屏幕取色】

'*************************************************
' 摘  要：全屏图片取色
' 作者: CBM666
' QQ：138449666
' 邮箱: samliu0812@163.com
' 转载请注明代码来源出处
' 本教材部份示例截图浏览地址：
' http://xiangce.baidu.com/picture/album/list/2853d800a41f08318a66aa5b12c232e7bcbc6455
' VIP 收费教学群: 203643302
' CBM666 免费 VB教学群: 44219538
' CBM666 免费 VB高级群: 120645325
'*************************************************

Option Explicit '强制必须声明变量
Private WithEvents Picture1 As PictureBox '自定义线上添加控件picture1的声明
Attribute Picture1.VB_VarHelpID = -1
Private WithEvents Picture2 As PictureBox '自定义线上添加控件picture2的声明
Attribute Picture2.VB_VarHelpID = -1
Dim MousePos As POINTAPI '定义MousePos对象
Private Sub Form_Load() '窗体载入事件
   AppDisk = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") '判断本地路径的\ 赋值给变量AppDisk
   Set Picture1 = Me.Controls.Add("VB.PictureBox", "Picture1") '线上添加Picture1控件
   Picture1.BorderStyle = 0 '图片框picture1设定为无边框
   Picture1.Visible = True '线上添加的控件默认为不可见 所以得加上这行让它 可见.
   Picture1.Move 5450, 7320, 6720, 495 '设定picture1的宽度与高度并移动到5450,7320的坐标位置
   Set Picture2 = Me.Controls.Add("VB.PictureBox", "Picture2") '线上添加Picture2控件
   Picture2.BorderStyle = 0 '图片框picture2设定为无边框
   Me.AutoRedraw = True '窗体自动重画为真
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2 '窗体居于屏幕中心位置
   Me.Picture = LoadPicture(AppDisk & "ColorSet.jpg") '本地路径下的ColorSet.jpg装载进窗体当背景图片
   WindowDC = GetWindowDC(0)  '获取屏幕的设备场景
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ReleaseDC 0, WindowDC '释放影像内存
   Controls.Remove ("Picture1") '移除动态添加的控件
   Controls.Remove ("Picture2")
   Set Form1 = Nothing '释放窗体占用内存
   End '退出结束程序
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ColorVal = GetPixel(Me.hdc, X \ 15, Y \ 15) '使用API GetPixel 获取颜色十进制值 坐标值XY单位是Twip缇 转换为像素Pixel需要除以15
   Call GetRGB '调用副程序GetRGB 返回RGB红绿蓝的颜色值
   '将获取到的颜色解析后的红绿蓝三个基本色在窗体标题栏显示
   Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '窗口标题显示颜色值
   Picture1.BackColor = RGB(Red, Green, Blue) '将右下角的Picture1 涂上当前鼠标指向的颜色同步刷新
   LabPos(1).Caption = CStr(Round(X / 15)) & "," & CStr(Round(Y / 15))
End Sub

Private Sub Command1_Click() '按钮点击事件
   On Error GoTo Errhandler ' 捕捉错误
   With CommonDialog1 '使用CommonDialog1控件对象
      .DialogTitle = "打开图片" '设置控件的标题名称
      .DefaultExt = ".jpg" ' 设置默认的扩展名
      .Filter = "所有支持的图片格式" & "(*.bmp;*.jpg;*.gif)|" & "*.bmp;*.jpg;*.gif)" '设定文件格式
      .ShowOpen ' 显示"文件打开"对话框
   End With
   Picture2.Picture = LoadPicture(CommonDialog1.FileName) '将选中的图片文件在Picture2载入显示
   Call DrawPicture '调用副程序 DrawPicture 画出载入的图像
Errhandler:
   If Err > 0 Then Exit Sub
End Sub

Sub DrawPicture()
   PicPhoto.Cls '清除PicPhoto图像
   '在PicPhoto内使用PicPhoto的PaintPicture方法在坐标100,100的地方画出宽度4000 高度2950 Picture2的图像
   PicPhoto.PaintPicture Picture2.Picture, 100, 100, 4000, 2950
End Sub

Private Sub Timer1_Timer()
   GetCursorPos MousePos
   LabPos(0).Caption = CStr(MousePos.X) & "," & CStr(MousePos.Y) '对应屏幕当前坐标XY值使用LabPos(0)标签显示
   '如果鼠标不再窗体内则清除 LabPos(1)标签的坐标值
   If MousePos.X < Me.Left / 15 Or MousePos.X > (Me.Left + Me.Width) / 15 Or MousePos.Y < Me.Top / 15 Or MousePos.Y > (Me.Top + Me.Height) / 15 Then
      LabPos(1).Caption = "" '清除 LabPos(1)标签的坐标值
   End If
   ColorVal = GetPixel(WindowDC, MousePos.X, MousePos.Y) '获取鼠标所指点的颜色
   Call GetRGB '调用颜色十进制值转换为RGB的副程序
   Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '窗口标题显示颜色值
   Picture1.BackColor = RGB(Red, Green, Blue) '将右下角的picture1 涂上当前鼠标指向的颜色同步刷新
End Sub

Private Sub Timer2_Timer()
   '****************************************** 检测鼠标左键是否被按下
   KeyResult = GetAsyncKeyState(1) '检测鼠标左键状态值并赋值给变量 KeyResult
   If KeyResult = -32767 Or KeyResult = -32768 Then '鼠标左键被按下了
      GetCursorPos MousePos
      ColorVal = GetPixel(WindowDC, MousePos.X, MousePos.Y)      '获取鼠标所指点的颜色
      Call GetRGB '调用颜色十进制值转换为RGB的副程序
      Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '窗口标题显示颜色值
      Picture1.BackColor = RGB(Red, Green, Blue) '将右下角的picture1 涂上当前鼠标指向的颜色同步刷新
      Text1(0).Text = CStr(ColorVal): Text1(1).Text = CStr(Red) & "," & CStr(Green) & "," & CStr(Blue): Text1(2).Text = "&H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00")
      'Clipboard.Clear
      Clipboard.SetText (Text1(2).Text)
   End If
   '****************************************** 检测鼠标右键是否被按下
   KeyResult = GetAsyncKeyState(2) '检测鼠标右键状态值并赋值给变量 KeyResult
   If KeyResult = -32767 Or KeyResult = -32768 Then '鼠标右键被按下了
      If Clipboard.GetFormat(vbCFBitmap) Then '如果剪贴板内容是图像
         Picture2.Picture = Clipboard.GetData '将剪贴板的内容贴进Picture2内
         Call DrawPicture '调用副程序 DrawPicture 画出图像
      End If
   End If
   '**************************************** 检测 Esc按键 是否被按下
   If GetAsyncKeyState(vbKeyEscape) Then Unload Me '按下了 Esc 键结束程序运行
End Sub

