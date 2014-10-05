Attribute VB_Name = "Module1"
Option Explicit '强制必须声明变量
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'定义全局变量
Public KeyResult&, Red&, Green&, Blue&, ColorVal&, WindowDC&, AppDisk$ '变量声明与型态定义 $=String文字型 &=Long长整型
Public Type POINTAPI '定义区间结构体
   X As Long
   Y As Long
End Type

'启动程序
Sub Main()
   Form1.Show '运行 Form1
End Sub

Public Sub GetRGB() '颜色十进制值转换为RGB的副程序
   '各别计算红绿蓝的RGB值 带入变量 Red Green Blue
   Red = (ColorVal And &HFF&)
   Green = (ColorVal And &HFF00&) \ 256
   Blue = (ColorVal And &HFF0000) \ 65536
End Sub
