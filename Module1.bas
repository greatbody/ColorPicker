Attribute VB_Name = "Module1"
Option Explicit 'ǿ�Ʊ�����������
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'����ȫ�ֱ���
Public KeyResult&, Red&, Green&, Blue&, ColorVal&, WindowDC&, AppDisk$ '������������̬���� $=String������ &=Long������
Public Type POINTAPI '��������ṹ��
   X As Long
   Y As Long
End Type

'��������
Sub Main()
   Form1.Show '���� Form1
End Sub

Public Sub GetRGB() '��ɫʮ����ֵת��ΪRGB�ĸ�����
   '��������������RGBֵ ������� Red Green Blue
   Red = (ColorVal And &HFF&)
   Green = (ColorVal And &HFF00&) \ 256
   Blue = (ColorVal And &HFF0000) \ 65536
End Sub
