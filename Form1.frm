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
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   "����������"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "16����ֵ:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "RGB ֵ:"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ʮ����ֵ:"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "ѡ��ͼƬ"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��������:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��Ļ����:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���������λ�á�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��2���������������ط���ȡͼƬ���Ƶ���������,Ȼ������ͼƬ��ʾ����ʹ������Ҽ���������Խ�ͼ�������ͼ����ʾ������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��1�������������Ļ���������Ҫ����ɫ,RGBֵ����ʾ�����桾�������������ı�����"
      BeginProperty Font 
         Name            =   "����"
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
'ժ¼�ԡ�CBM666 VB���ʾ���̲� ͼ��ƪ_��Ļȡɫ��

'*************************************************
' ժ  Ҫ��ȫ��ͼƬȡɫ
' ����: CBM666
' QQ��138449666
' ����: samliu0812@163.com
' ת����ע��������Դ����
' ���̲Ĳ���ʾ����ͼ�����ַ��
' http://xiangce.baidu.com/picture/album/list/2853d800a41f08318a66aa5b12c232e7bcbc6455
' VIP �շѽ�ѧȺ: 203643302
' CBM666 ��� VB��ѧȺ: 44219538
' CBM666 ��� VB�߼�Ⱥ: 120645325
'*************************************************

Option Explicit 'ǿ�Ʊ�����������
Private WithEvents Picture1 As PictureBox '�Զ���������ӿؼ�picture1������
Attribute Picture1.VB_VarHelpID = -1
Private WithEvents Picture2 As PictureBox '�Զ���������ӿؼ�picture2������
Attribute Picture2.VB_VarHelpID = -1
Dim MousePos As POINTAPI '����MousePos����
Private Sub Form_Load() '���������¼�
   AppDisk = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") '�жϱ���·����\ ��ֵ������AppDisk
   Set Picture1 = Me.Controls.Add("VB.PictureBox", "Picture1") '�������Picture1�ؼ�
   Picture1.BorderStyle = 0 'ͼƬ��picture1�趨Ϊ�ޱ߿�
   Picture1.Visible = True '������ӵĿؼ�Ĭ��Ϊ���ɼ� ���Եü����������� �ɼ�.
   Picture1.Move 5450, 7320, 6720, 495 '�趨picture1�Ŀ����߶Ȳ��ƶ���5450,7320������λ��
   Set Picture2 = Me.Controls.Add("VB.PictureBox", "Picture2") '�������Picture2�ؼ�
   Picture2.BorderStyle = 0 'ͼƬ��picture2�趨Ϊ�ޱ߿�
   Me.AutoRedraw = True '�����Զ��ػ�Ϊ��
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2 '���������Ļ����λ��
   Me.Picture = LoadPicture(AppDisk & "ColorSet.jpg") '����·���µ�ColorSet.jpgװ�ؽ����嵱����ͼƬ
   WindowDC = GetWindowDC(0)  '��ȡ��Ļ���豸����
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ReleaseDC 0, WindowDC '�ͷ�Ӱ���ڴ�
   Controls.Remove ("Picture1") '�Ƴ���̬��ӵĿؼ�
   Controls.Remove ("Picture2")
   Set Form1 = Nothing '�ͷŴ���ռ���ڴ�
   End '�˳���������
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ColorVal = GetPixel(Me.hdc, X \ 15, Y \ 15) 'ʹ��API GetPixel ��ȡ��ɫʮ����ֵ ����ֵXY��λ��Twip� ת��Ϊ����Pixel��Ҫ����15
   Call GetRGB '���ø�����GetRGB ����RGB����������ɫֵ
   '����ȡ������ɫ������ĺ�������������ɫ�ڴ����������ʾ
   Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '���ڱ�����ʾ��ɫֵ
   Picture1.BackColor = RGB(Red, Green, Blue) '�����½ǵ�Picture1 Ϳ�ϵ�ǰ���ָ�����ɫͬ��ˢ��
   LabPos(1).Caption = CStr(Round(X / 15)) & "," & CStr(Round(Y / 15))
End Sub

Private Sub Command1_Click() '��ť����¼�
   On Error GoTo Errhandler ' ��׽����
   With CommonDialog1 'ʹ��CommonDialog1�ؼ�����
      .DialogTitle = "��ͼƬ" '���ÿؼ��ı�������
      .DefaultExt = ".jpg" ' ����Ĭ�ϵ���չ��
      .Filter = "����֧�ֵ�ͼƬ��ʽ" & "(*.bmp;*.jpg;*.gif)|" & "*.bmp;*.jpg;*.gif)" '�趨�ļ���ʽ
      .ShowOpen ' ��ʾ"�ļ���"�Ի���
   End With
   Picture2.Picture = LoadPicture(CommonDialog1.FileName) '��ѡ�е�ͼƬ�ļ���Picture2������ʾ
   Call DrawPicture '���ø����� DrawPicture ���������ͼ��
Errhandler:
   If Err > 0 Then Exit Sub
End Sub

Sub DrawPicture()
   PicPhoto.Cls '���PicPhotoͼ��
   '��PicPhoto��ʹ��PicPhoto��PaintPicture����������100,100�ĵط��������4000 �߶�2950 Picture2��ͼ��
   PicPhoto.PaintPicture Picture2.Picture, 100, 100, 4000, 2950
End Sub

Private Sub Timer1_Timer()
   GetCursorPos MousePos
   LabPos(0).Caption = CStr(MousePos.X) & "," & CStr(MousePos.Y) '��Ӧ��Ļ��ǰ����XYֵʹ��LabPos(0)��ǩ��ʾ
   '�����겻�ٴ���������� LabPos(1)��ǩ������ֵ
   If MousePos.X < Me.Left / 15 Or MousePos.X > (Me.Left + Me.Width) / 15 Or MousePos.Y < Me.Top / 15 Or MousePos.Y > (Me.Top + Me.Height) / 15 Then
      LabPos(1).Caption = "" '��� LabPos(1)��ǩ������ֵ
   End If
   ColorVal = GetPixel(WindowDC, MousePos.X, MousePos.Y) '��ȡ�����ָ�����ɫ
   Call GetRGB '������ɫʮ����ֵת��ΪRGB�ĸ�����
   Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '���ڱ�����ʾ��ɫֵ
   Picture1.BackColor = RGB(Red, Green, Blue) '�����½ǵ�picture1 Ϳ�ϵ�ǰ���ָ�����ɫͬ��ˢ��
End Sub

Private Sub Timer2_Timer()
   '****************************************** ����������Ƿ񱻰���
   KeyResult = GetAsyncKeyState(1) '���������״ֵ̬����ֵ������ KeyResult
   If KeyResult = -32767 Or KeyResult = -32768 Then '��������������
      GetCursorPos MousePos
      ColorVal = GetPixel(WindowDC, MousePos.X, MousePos.Y)      '��ȡ�����ָ�����ɫ
      Call GetRGB '������ɫʮ����ֵת��ΪRGB�ĸ�����
      Me.Caption = CStr(ColorVal) & "--- R:" & CStr(Red) & ",G:" & CStr(Green) & ",B:" & CStr(Blue) & "--- H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00") '���ڱ�����ʾ��ɫֵ
      Picture1.BackColor = RGB(Red, Green, Blue) '�����½ǵ�picture1 Ϳ�ϵ�ǰ���ָ�����ɫͬ��ˢ��
      Text1(0).Text = CStr(ColorVal): Text1(1).Text = CStr(Red) & "," & CStr(Green) & "," & CStr(Blue): Text1(2).Text = "&H" & Format(Hex(Blue), "00") & Format(Hex(Green), "00") & Format(Hex(Red), "00")
      'Clipboard.Clear
      Clipboard.SetText (Text1(2).Text)
   End If
   '****************************************** �������Ҽ��Ƿ񱻰���
   KeyResult = GetAsyncKeyState(2) '�������Ҽ�״ֵ̬����ֵ������ KeyResult
   If KeyResult = -32767 Or KeyResult = -32768 Then '����Ҽ���������
      If Clipboard.GetFormat(vbCFBitmap) Then '���������������ͼ��
         Picture2.Picture = Clipboard.GetData '�����������������Picture2��
         Call DrawPicture '���ø����� DrawPicture ����ͼ��
      End If
   End If
   '**************************************** ��� Esc���� �Ƿ񱻰���
   If GetAsyncKeyState(vbKeyEscape) Then Unload Me '������ Esc ��������������
End Sub

