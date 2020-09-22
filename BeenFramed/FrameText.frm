VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrameText 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy Frame Text"
   ClientHeight    =   5745
   ClientLeft      =   2940
   ClientTop       =   1485
   ClientWidth     =   8805
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FrameText.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "FrameText.frx":030A
      Top             =   5160
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog cmOpenFile 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd3 
      Caption         =   "Save File"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Press this to save the text above as a web page!"
      Top             =   5210
      Width           =   975
   End
   Begin VB.TextBox Txt2 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Text            =   "Frame"
      ToolTipText     =   "This is the name of the page, you can change it, just remeber that you can save it automatically with the extension .html"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "ClipBoard"
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      ToolTipText     =   "Press here to copy the text to memory"
      Top             =   5210
      Width           =   855
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      ToolTipText     =   "Press here to go back to the frame form"
      Top             =   5210
      Width           =   855
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      HideSelection   =   0   'False
      Left            =   10
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "FrameText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd1_Click()
FrameText.Hide
Frame.Show
End Sub

Private Sub Cmd2_Click()
On Error GoTo Err_Handler
Clipboard.Clear   ' Clear Clipboard.
If Txt.SelLength >= 1 Then
Clipboard.SetText Txt.SelText, 1   ' Put text on Clipboard
Else
Clipboard.SetText Txt.Text, 1   ' Put text on Clipboard
End If
Exit Sub

Err_Handler:

End Sub

Private Sub Cmd3_Click()
On Error GoTo Save_Error
If Txt.Text <> "" Then
Dim intFile As Integer
With cmOpenFile
    .FileName = Txt2.Text
    .Filter = "HTML Files (*.html, *.htm, *.asp, *.shtml, *.css) | *.html; *htm; *.asp; *.shtml; *.css; |Text Files (*.txt) | *.txt; |All Files (*.*) | *.*"
    .DialogTitle = "Save Your Web Page File?"
    cmOpenFile.Flags = &H2
    cmOpenFile.ShowSave
End With
    intFile = FreeFile
    ' open will overwrite an existing file for output
    Open cmOpenFile.FileName For Output As intFile
    ' put the whole text box into the file
    Print #intFile, Txt.Text
    ' now update the filename on the caption
        FrameText.Caption = "Copy Frame Text " & cmOpenFile.FileName & " Has Been Saved"
        Close #intFile
    ' let the program know the file has been saved
    End If
    Exit Sub
Save_Error:

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Public Static Sub Form_Load()
FrameText.Left = (Screen.Width / 2) - (FrameText.Width / 2)
FrameText.Top = (Screen.Height / 2) - (FrameText.Height / 2)

End Sub
