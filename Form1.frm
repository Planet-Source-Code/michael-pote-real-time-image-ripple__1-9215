VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Waverizer V2.0 by MTECH Designs."
   ClientHeight    =   8175
   ClientLeft      =   2085
   ClientTop       =   270
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   Begin VB.CommandButton Command9 
      Caption         =   "+"
      Height          =   315
      Left            =   2835
      TabIndex        =   26
      ToolTipText     =   "Add a Preset"
      Top             =   5970
      Width           =   330
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   840
      TabIndex        =   25
      Text            =   "Presets"
      Top             =   5970
      Width           =   1980
   End
   Begin VB.Timer Timer1 
      Left            =   -330
      Top             =   5910
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Run"
      Height          =   330
      Left            =   3300
      TabIndex        =   23
      Top             =   5940
      Width           =   945
   End
   Begin VB.OptionButton Option6 
      Caption         =   "Ripple"
      Height          =   225
      Left            =   5250
      TabIndex        =   22
      Top             =   7440
      Width           =   1290
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ripple"
      Height          =   360
      Left            =   540
      TabIndex        =   21
      Top             =   3375
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   240
      Left            =   135
      Max             =   20
      Min             =   1
      TabIndex        =   19
      Top             =   7920
      Value           =   1
      Width           =   6255
   End
   Begin VB.OptionButton Option5 
      Caption         =   "2Tap Wave"
      Height          =   300
      Left            =   4065
      TabIndex        =   18
      Top             =   7395
      Width           =   1290
   End
   Begin VB.CommandButton Command6 
      Caption         =   "2Tap Wave"
      Height          =   360
      Left            =   555
      TabIndex        =   17
      Top             =   3780
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open a Picture"
      Height          =   345
      Left            =   4290
      TabIndex        =   16
      Top             =   5925
      Width           =   2100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Switching station"
      Height          =   285
      Left            =   2295
      TabIndex        =   15
      Top             =   3015
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Cosine wave"
      Height          =   300
      Left            =   2775
      TabIndex        =   14
      Top             =   7395
      Width           =   1290
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cosine Wave"
      Height          =   360
      Left            =   540
      TabIndex        =   13
      Top             =   4185
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Both"
      Height          =   360
      Left            =   570
      TabIndex        =   12
      Top             =   4605
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Horizontal"
      Height          =   360
      Left            =   585
      TabIndex        =   11
      Top             =   5025
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Both"
      Height          =   240
      Left            =   2025
      TabIndex        =   10
      Top             =   7425
      Width           =   1035
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Verticle"
      Height          =   240
      Left            =   1110
      TabIndex        =   9
      Top             =   7425
      Width           =   1035
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Horizontal"
      Height          =   240
      Left            =   45
      TabIndex        =   8
      Top             =   7425
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   285
      Left            =   135
      Max             =   100
      Min             =   -100
      TabIndex        =   6
      Top             =   6975
      Value           =   10
      Width           =   4815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   285
      Left            =   135
      Max             =   15
      Min             =   1
      TabIndex        =   4
      Top             =   6660
      Value           =   1
      Width           =   4830
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   135
      Max             =   1080
      TabIndex        =   3
      Top             =   6315
      Width           =   6285
   End
   Begin VB.CommandButton Command 
      Caption         =   "Verticle"
      Height          =   360
      Left            =   585
      TabIndex        =   2
      Top             =   5445
      Visible         =   0   'False
      Width           =   5445
   End
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   6510
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   419
      TabIndex        =   1
      Top             =   1230
      Visible         =   0   'False
      Width           =   6285
   End
   Begin VB.PictureBox picDest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5955
      Left            =   30
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   419
      TabIndex        =   0
      Top             =   -45
      Width           =   6345
   End
   Begin VB.Label Label3 
      Caption         =   "Presets:"
      Height          =   300
      Left            =   195
      TabIndex        =   24
      Top             =   6015
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Quality ---------------------------------------------------Medium------------------------------------------------   Speed"
      Height          =   195
      Left            =   165
      TabIndex        =   20
      Top             =   7710
      Width           =   6135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude: 10"
      Height          =   195
      Left            =   4995
      TabIndex        =   7
      Top             =   7020
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Waves: 1"
      Height          =   195
      Left            =   5010
      TabIndex        =   5
      Top             =   6705
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public X As Long, Y As Long, Ang As Long, Dep As Long, Mult As Long, Pi As Variant, Wave As Variant, Heig As Long, Wid As Long, Opt As Integer, Ang2 As Long, Wave2, S As Long
Public PreCnt As Long
Private Dist() As Variant
'WAVERIZER by MTECH Designs (Michael Pote)
'-----------------------------------------
'Only 7 Lines of code to make a big picture occilate!
'Mult - Number of waves
'Dep - Amplitude / Depth
'S - Number of Steps used for quality control

Private Function Distance(sx, sy, Ex, Ey) As Long
Distance = Sqr((Ex - sx) ^ 2 + (Ey - sy) ^ 2)
End Function

Private Sub Combo1_Click()
Dim Path As String
Path = App.Path & "\Presets.ini"
Opt = ReadINI(Path, "Preset" & Combo1.ListIndex, "Opt")
Mult = ReadINI(Path, "Preset" & Combo1.ListIndex, "Mult")
Dep = ReadINI(Path, "Preset" & Combo1.ListIndex, "Dep")
S = ReadINI(Path, "Preset" & Combo1.ListIndex, "S")
UpdateSliders
End Sub

Private Sub Command1_Click()
'Choose which wave to display
Select Case Opt
Case 1
Command_Click
Case 2
Command2_Click
Case 3
Command3_Click
Case 4
Command4_Click
Case 5
Command6_Click
Case 6
Command7_Click
End Select
End Sub

Private Sub Command_Click()
For X = 1 To Heig
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = Sin(Pi * Ang) * Dep
BitBlt picDest.hdc, Wave, X, Wid, 2, picSrc.hdc, 0, X, SRCCOPY
Next
picDest.Refresh
End Sub

Private Sub Command2_Click()
For X = 1 To Wid
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = Sin(Pi * Ang) * Dep
BitBlt picDest.hdc, X, Wave, 2, Heig, picSrc.hdc, X, 0, SRCCOPY
Next
picDest.Refresh

End Sub

Private Sub Command3_Click()
For X = 1 To Heig
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = Sin(Pi * Ang) * Dep
BitBlt picDest.hdc, Wave, X + Wave, Wid, 4, picSrc.hdc, 0, X, SRCCOPY
Next
picDest.Refresh

End Sub

Private Sub Command4_Click()
For X = 1 To Heig
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = 1 / Sin(Pi * Ang) * Dep
BitBlt picDest.hdc, Wave, X, Wid, 4, picSrc.hdc, 0, X, SRCCOPY
Next
picDest.Refresh
End Sub

Private Sub Command5_Click()
Dim File As String
File = OpenDialog(Form1, "Pictures|*.bmp|Jpegs|*.jpg", "Open a picture", "")
If File = "" Then Exit Sub
picSrc.Picture = LoadPicture(File)
picDest.Width = picSrc.Width
picDest.Height = picSrc.Height
DoTable
picDest.Move (ScaleWidth / 2) - (Wid / 2), (400 / 2) - (Heig / 2)
End Sub

Private Sub Command6_Click()
For X = 1 To Wid Step S
DoEvents
Ang = HScroll1.Value + (X * Mult)
Wave = Sin(Pi * Ang) * Dep
For Y = 1 To Heig Step S
Ang2 = HScroll1.Value + (Y * Mult)
Wave2 = Cos(Pi * Ang2) * Dep
BitBlt picDest.hdc, X + Wave2, Y + Wave, S, S, picSrc.hdc, X, Y, SRCCOPY
Next
Next
picDest.Refresh
End Sub

Private Sub Command7_Click()
For X = 1 To Wid Step S
DoEvents
For Y = 1 To Heig Step S
Ang = HScroll1.Value + (Dist(X, Y) * Mult)
Wave = Cos(Pi * Ang) * Dep
BitBlt picDest.hdc, X + Wave, Y + Wave, S + 1, S + 1, picSrc.hdc, X, Y, SRCCOPY
Next
Next
picDest.Refresh
End Sub


Private Sub Command8_Click()
If Command8.Caption = "Run" Then
Timer1.Interval = 1
Command8.Caption = "Stop"
Else
Command8.Caption = "Run"
End If
End Sub

Private Sub Command9_Click()
Dim Name As String, Path As String
Name = InputBox("Type a name for this preset:", "Name", "Untitled")
Path = App.Path & "\Presets.ini"
WriteINI Path, "General", "Num", CStr(Combo1.ListCount)
WriteINI Path, "Preset" & Combo1.ListCount, "Name", Name
WriteINI Path, "Preset" & Combo1.ListCount, "Mult", CStr(Mult)
WriteINI Path, "Preset" & Combo1.ListCount, "Dep", CStr(Dep)
WriteINI Path, "Preset" & Combo1.ListCount, "Opt", CStr(Opt)
WriteINI Path, "Preset" & Combo1.ListCount, "S", CStr(S)
Combo1.AddItem Name
End Sub

Private Sub Form_Load()
Dep = 30
Pi = (3.1456 / 180)
Mult = 1
Opt = 2
S = 1
DoTable
LoadPresets
End Sub
Sub LoadPresets()
If Dir(App.Path & "\Presets.ini") = "" Then Exit Sub
PreCnt = CInt(ReadINI(App.Path & "\Presets.ini", "General", "Num"))
For i = 0 To PreCnt
Combo1.AddItem ReadINI(App.Path & "\Presets.ini", "Preset" & i, "Name")
Next
End Sub
Sub DoTable()
Heig = picDest.ScaleHeight
Wid = picDest.ScaleWidth
ReDim Dist(1 To Wid, 1 To Heig) As Variant
For X = 1 To Wid
For Y = 1 To Heig
Dist(X, Y) = Distance(Wid / 2, Heig / 2, X, Y)
Next
Next
End Sub

Private Sub HScroll1_Scroll()
Command1_Click

End Sub

Private Sub HScroll2_Change()
Mult = HScroll2.Value
Command1_Click
Label1.Caption = "Number of Waves: " & Mult
End Sub

Private Sub HScroll2_Scroll()
HScroll2_Change
End Sub

Private Sub HScroll3_Change()
HScroll3_Scroll
End Sub

Private Sub HScroll3_Scroll()
Dep = HScroll3.Value
Label2.Caption = "Amplitude: " & HScroll3.Value
Command1_Click
End Sub

Private Sub HScroll4_Scroll()
S = HScroll4.Value
Command1_Click
End Sub

Private Sub Option1_Click()
Opt = 2
HScroll4.Enabled = False
End Sub

Private Sub Option2_Click()
Opt = 1
HScroll4.Enabled = False
End Sub

Private Sub Option3_Click()
Opt = 3
HScroll4.Enabled = False
End Sub

Private Sub Option4_Click()
Opt = 4
HScroll4.Enabled = False
End Sub

Private Sub Option5_Click()
Opt = 5
HScroll4.Enabled = True
End Sub

Private Sub Option6_Click()
Opt = 6
HScroll4.Enabled = True
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = 0
Do
DoEvents
If HScroll1.Value >= HScroll1.Max - 5 Then HScroll1.Value = HScroll1.Min
On Error Resume Next
HScroll1.Value = HScroll1.Value + 5
Command1_Click
Loop While Command8.Caption = "Stop"
End Sub

Sub UpdateSliders()
Select Case Opt
Case 1
Option1.Value = True
Case 2
Option2.Value = True
Case 3
Option3.Value = True
Case 4
Option4.Value = True
Case 5
Option5.Value = True
Case 6
Option6.Value = True
End Select
HScroll2.Value = Mult
HScroll3.Value = Dep
HScroll4.Value = S
Command1_Click
End Sub
