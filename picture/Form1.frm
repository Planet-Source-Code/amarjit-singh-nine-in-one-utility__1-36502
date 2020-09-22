VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   468
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   3165
      Left            =   9000
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3105
      ScaleWidth      =   7440
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   3240
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   12
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Timer Timer8 
      Interval        =   100
      Left            =   4200
      Top             =   3240
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2055
      Left            =   3000
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   11
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   2160
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      Picture         =   "Form1.frx":271A
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   2160
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   3840
      Top             =   2160
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2055
      Left            =   2940
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   7
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   2160
   End
   Begin VB.Timer Timer2 
      Interval        =   150
      Left            =   2160
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2160
      Top             =   2160
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   2055
      Left            =   5760
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   4800
      Width           =   2535
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amarjit Singh"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Coded"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   120
      Picture         =   "Form1.frx":39C8
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   4800
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2055
      Left            =   2940
      Picture         =   "Form1.frx":4519
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      Height          =   2055
      Left            =   120
      Picture         =   "Form1.frx":174A3
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim w1 As Integer, h1 As Integer, x1 As Integer
Dim dx1 As Integer, y1 As Integer, y2 As Integer, dx2 As Integer
Dim x3 As Integer, y3 As Integer, dx3 As Integer, dy3 As Integer
Dim y41 As Integer, y42 As Integer, y43 As Integer
Dim st As Long, en As Long
Dim r1 As Integer, g1 As Integer, b1 As Integer
Dim r2 As Integer, g2 As Integer, b2 As Integer
Dim dr As Single, dg As Single, db As Single
Dim i As Integer, flag As Integer
Dim h As Integer, m As Integer, s As Integer
Dim xx As Integer

Private Type pos
x As Integer
y As Integer
dy As Integer
End Type


Private Type SFONT
 size As Integer
 str As String * 1
 x As Integer
 y As Integer
 dy As Integer
 cl As Integer
 End Type

Dim t1(50) As SFONT
Dim t(50) As pos

Dim ii As Integer, jj As Integer, flag1 As Integer, kk As Integer

Private Sub Form_Load()
w1 = Picture1.ScaleWidth
h1 = Picture1.ScaleHeight
If w1 > h1 Then
x1 = h1 / 2
Else
x1 = w1 / 2
End If

pic1
pic2
pic3
pic4
pic5
pic6
pic7
pic8
pic9
End Sub
Private Sub pic1()
y1 = x1
dx1 = 3

Timer1.Enabled = True
End Sub

Private Sub pic2()
Timer2.Enabled = True

y2 = 6
dx2 = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
End Sub

Private Sub Timer1_Timer()
Picture1.Cls
Picture1.FillStyle = 0
Picture1.FillColor = RGB(255, 164, 101)
Picture1.Circle (w1 / 2, h1 / 2), y1, RGB(255, 164, 101)
If y1 > x1 Or y1 < 2 Then dx1 = -dx1
y1 = y1 + dx1

End Sub

Private Sub Timer2_Timer()
Picture2.Cls
Dim i As Integer
Picture2.FillStyle = 0
Picture2.FillColor = vbBlue
'For i = -y2 To y2 Step 4
Picture2.Line (w1 / 2 - y2, h1 / 2 - y2)-(w1 / 2 + y2, h1 / 2 + y2), vbBlue, B
'Next
If y2 > x1 Or y2 < 2 Then dx2 = -dx2
y2 = y2 + dx2
End Sub

Private Sub pic3()
x3 = 20
y3 = 20
dx3 = 5
dy3 = 5

Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Picture3.Cls
Picture3.FillStyle = 0
'Picture3.FillColor = Picture3.BackColor
'Picture3.Circle (x3, y3), 16
x3 = x3 + dx3
y3 = y3 + dy3
Picture3.FillColor = vbCyan
 Picture3.Circle (x3, y3), 5, vbCyan
If x3 > w1 - 5 Or x3 < 5 Then dx3 = -dx3: Beep
If y3 > h1 - 5 Or y3 < 5 Then dy3 = -dy3: Beep
End Sub

Private Sub pic4()
Timer4.Enabled = True
y43 = h1 - Label1.Height
y42 = h1 - Label1.Height - Label2.Height
y41 = h1 - Label1.Height - Label2.Height - Label4.Height

End Sub
Private Sub Timer4_Timer()
Label1.Top = y41
y41 = y41 - 2
If y41 < -10 Then y41 = h1
Label2.Top = y42
y42 = y42 - 2
If y42 < -10 Then y42 = h1

Label4.Top = y43

y43 = y43 - 2
If y43 < -10 Then y43 = h1

End Sub

Private Sub pic5()
st = RGB(255, 0, 0)
en = RGB(0, 255, 0)
r1 = getred(st)
g1 = getgreen(st)
b1 = getblue(st)
r2 = getred(en)
g2 = getgreen(en)
b2 = getblue(en)
dr = (r2 - r1) / w1
dg = (g2 - g1) / w1
db = (b2 - b1) / w1
i = 0
flag = 0
Timer5.Enabled = True
End Sub
Private Sub Timer5_Timer()
Dim j As Integer
If flag = 0 Then
For j = 0 To 5
Picture5.Line (i, 0)-(i, h1), RGB(r1 + i * dr, g1 + i * dg, b1 + i * db)
i = i + 1
If i > w1 Then flag = 1: Exit For
Next
Else
For j = 0 To 5
i = i - 1
If i < 0 Then flag = 0: i = 0: Exit For
Picture5.Line (i, 0)-(i, h1), RGB(0, 0, 0)

Next
End If
'For i = w1 To 0 Step -1
'Picture5.Line (i, 0)-(i, h1), Picture5.BackColor
'Next
End Sub
Private Sub pic6()
If w1 > h1 Then
xx = h1 / 2
Else
xx = w1 / 2
End If
Timer6.Enabled = True
End Sub

'Private Sub Timer6_Timer()
'Picture6.Cls
'Picture6.FontBold = True
'Picture6.FontSize = Rnd() * 16 + 1
'Picture6.CurrentY = v
'Picture6.Print "A"
'v = v + 16
'If v > h1 Then v = 0
'End Sub
Private Sub Timer6_Timer()
Picture6.Cls
Dim i As Integer
Dim st As String
For i = 0 To 360 Step 6
If i Mod 30 = 0 Then
Picture6.Line (w1 / 2 + (xx - 16) * Sin(i * 3.14 / 180), h1 / 2 - (xx - 16) * Cos(i * 3.14 / 180))-(w1 / 2 + (xx - 1) * Sin(i * 3.14 / 180), h1 / 2 - (xx - 1) * Cos(i * 3.14 / 180)), RGB(255, 0, 0)
Else
Picture6.Line (w1 / 2 + (xx - 8) * Sin(i * 3.14 / 180), h1 / 2 - (xx - 8) * Cos(i * 3.14 / 180))-(w1 / 2 + (xx - 1) * Sin(i * 3.14 / 180), h1 / 2 - (xx - 1) * Cos(i * 3.14 / 180)), RGB(255, 255, 255)
End If
Next
h = Hour(Now)
m = Minute(Now)
s = Second(Now)
If s Mod 2 = 0 Then
st = h & " " & m & " " & s
Else
 st = h & ":" & m & ":" & s
 End If
Label3.Caption = st
Picture6.DrawWidth = 6
Picture6.Circle (w1 / 2, h1 / 2), xx - 2, vbCyan
Picture6.Circle (w1 / 2, h1 / 2), 3, vbRed
Picture6.DrawWidth = 2
Picture6.Line (w1 / 2, h1 / 2)-(w1 / 2 + (xx - 12) * Sin(s * 6 * 3.14 / 180), h1 / 2 - (xx - 12) * Cos(s * 6 * 3.14 / 180)), vbMagenta
Picture6.Line (w1 / 2, h1 / 2)-(w1 / 2 + (xx - 20) * Sin(m * 6 * 3.14 / 180), h1 / 2 - (xx - 20) * Cos(m * 6 * 3.14 / 180)), vbGreen
Picture6.Line (w1 / 2, h1 / 2)-(w1 / 2 + (xx - 30) * Sin(h * 30 * 3.14 / 180), h1 / 2 - (xx - 30) * Cos(h * 30 * 3.14 / 180)), vbBlue
End Sub
Private Sub pic7()
Dim i As Integer
For i = 0 To UBound(t)
t(i).x = Rnd() * w1 + 1
t(i).y = Rnd() * 16
t(i).dy = Rnd() * 15 + 1
Next
Picture7.FillColor = vbWhite
Picture7.FillStyle = 0
Timer7.Enabled = True
End Sub

Private Sub Timer7_Timer()
Dim i As Integer
Picture7.Cls
For i = 0 To UBound(t)
Picture7.Circle (t(i).x, t(i).y), 1, vbWhite
t(i).y = t(i).y + t(i).dy
If t(i).y > h1 Then t(i).x = Rnd() * w1: t(i).y = 7
Next
End Sub
Private Sub pic8()
Dim i As Integer
For i = 0 To UBound(t1)
t1(i).size = Rnd() * 10 + 10
t1(i).str = Rnd() * 8 + 1
t1(i).x = Rnd() * w1
t1(i).y = Rnd() * 7
t1(i).dy = Rnd() * 15 + 1
t1(i).cl = Rnd() * 14 + 1
Next
Picture8.FontBold = True
Timer8.Enabled = True
End Sub

Private Sub Timer8_Timer()
Picture8.Cls
Dim i As Integer
For i = 0 To UBound(t1)
Picture8.ForeColor = QBColor(t1(i).cl)
Picture8.CurrentX = t1(i).x
Picture8.CurrentY = t1(i).y
Picture8.FontSize = t1(i).size
Picture8.Print t1(i).str
t1(i).y = t1(i).y + t1(i).dy
If t1(i).y > h1 Then t1(i).y = Rnd() * 7: t1(i).x = Rnd() * w1: t1(i).dy = Rnd() * 14 + 1
Next
End Sub
Private Sub pic9()
ii = 0
jj = 0
kk = 0
flag1 = 0
Timer9.Enabled = True
End Sub

Private Sub Timer9_Timer()
Picture9.Cls
'On Error Resume Next
Picture9.PaintPicture pic.Picture, kk - 64, Picture9.ScaleHeight - 70, 128, 68, ii * 128, jj * 128, 130, 68
kk = kk + Picture9.ScaleWidth / 12
ii = ii + 1
If ii = 4 Then jj = 1: ii = 0: flag1 = 1
If flag1 = 1 And ii = 2 Then jj = 0: ii = 0: flag1 = 0
If kk > Picture9.ScaleWidth Then kk = -64
End Sub
