VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Color 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pick a Kouler not you Nose"
   ClientHeight    =   4920
   ClientLeft      =   5190
   ClientTop       =   1485
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Icon            =   "color.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   1620
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   960
      Max             =   255
      SmallChange     =   5
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   960
      Max             =   255
      SmallChange     =   5
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   960
      Max             =   255
      SmallChange     =   5
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.CommandButton CmdButton3 
      Caption         =   "Use the R G B"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Color3"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Color1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmOpenFile 
      Left            =   5040
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "255 Red"
      Height          =   255
      Left            =   3480
      TabIndex        =   22
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   21
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "255 Blue"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "255 Green"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Blue 0"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Green 0"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Red 0"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label7 
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private HSval1 As Byte
Private HSval2 As Byte
Private HSval3 As Byte

Private Function MySub(MyColor As Long)
Dim TmpStr As String * 6, NewClr As String, I As Integer, Letter As String * 1, NumStr As Integer, NewStr As String * 1
TempNumber = MyColor
TmpStr = Format(Hex(MyColor), "@@@@@@")
For I = 1 To 6
Letter = Mid(TmpStr, I, 1)
NumStr = Asc(Letter)
If NumStr = 32 Then NumStr = 48
NewStr = Chr(NumStr)
NewClr = NewClr + NewStr
Next I
RStr = Right(NewClr, 2)
LStr = Left(NewClr, 2)
MStr = Mid(NewClr, 3, 2)
MySub = RStr & MStr & LStr

End Function

Private Sub cmdButton_Click()
Dim cdCCFullOpen As Long
cmOpenFile.CancelError = True
cmOpenFile.Flags = cdCCFullOpen
On Error GoTo No_Color  ' traps error to stop program crash!
cmOpenFile.ShowColor
Label1.BackColor = cmOpenFile.Color
Label2.Caption = MySub(Label1.BackColor)

Exit Sub
No_Color:
End Sub

Private Sub cmdButton2_Click()
Dim cdCCFullOpen As Long
cmOpenFile.CancelError = True
cmOpenFile.Flags = cdCCFullOpen
On Error GoTo No_Color  ' traps error to stop program crash!
cmOpenFile.ShowColor
Label3.BackColor = cmOpenFile.Color
Label5.Caption = MySub(Label3.BackColor)
Exit Sub
No_Color:
End Sub
Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText Label2.Caption
End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText Label5.Caption
End Sub

Private Sub Command3_Click()
Clipboard.Clear
Clipboard.SetText Label17.Caption
End Sub

Private Sub CmdButton3_Click()
On Error GoTo No_Color  ' traps error to stop program crash!
Label7.BackColor = RGB(HSval1, HSval2, HSval3)
Label17.Caption = MySub(Label7.BackColor)
Exit Sub
No_Color:
End Sub

Private Sub Command4_Click()
Unload Me
Frame.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
Color.Left = (Screen.Width / 2) - (Color.Width / 2)
Color.Top = (Screen.Height / 2) - (Color.Height / 2)
HScroll1.Value = 0
HScroll2.Value = 0
HScroll3.Value = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Frame.Show
End Sub

Private Sub HScroll1_Change()
Label14.Caption = HScroll1.Value
HSval1 = HScroll1.Value
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label17.Caption = MySub(Label7.BackColor)
End Sub

Private Sub HScroll2_Change()
Label15.Caption = HScroll2.Value
HSval2 = HScroll2.Value
HScroll2.Value = HSval2
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label17.Caption = MySub(Label7.BackColor)
End Sub

Private Sub HScroll3_Change()
Label16.Caption = HScroll3.Value
HSval3 = HScroll3.Value
HScroll3.Value = HSval3
Label7.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
Label17.Caption = MySub(Label7.BackColor)
End Sub
