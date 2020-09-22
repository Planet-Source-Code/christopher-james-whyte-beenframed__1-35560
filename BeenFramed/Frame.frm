VERSION 5.00
Begin VB.Form Frame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I've Been Framed"
   ClientHeight    =   8145
   ClientLeft      =   3675
   ClientTop       =   2025
   ClientWidth     =   8025
   Icon            =   "Frame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd6 
      Caption         =   "Show Tips"
      Height          =   375
      Left            =   6960
      TabIndex        =   227
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Cmd4 
      Caption         =   "About"
      Height          =   375
      Left            =   6960
      TabIndex        =   226
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   2235
      Locked          =   -1  'True
      TabIndex        =   225
      Text            =   "5"
      Top             =   5070
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   200
      Locked          =   -1  'True
      TabIndex        =   224
      Text            =   "1"
      Top             =   3460
      Width           =   75
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   380
      Locked          =   -1  'True
      TabIndex        =   223
      Text            =   "2"
      Top             =   3460
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   222
      Text            =   "3"
      Top             =   3460
      Width           =   60
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   6440
      Locked          =   -1  'True
      TabIndex        =   221
      Text            =   "6"
      Top             =   7440
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   6440
      Locked          =   -1  'True
      TabIndex        =   220
      Text            =   "5"
      Top             =   6990
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   5690
      Locked          =   -1  'True
      TabIndex        =   219
      Text            =   "5"
      Top             =   7490
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   5690
      Locked          =   -1  'True
      TabIndex        =   218
      Text            =   "4"
      Top             =   6120
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   217
      Text            =   "5"
      Top             =   7710
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   2220
      Locked          =   -1  'True
      TabIndex        =   216
      Text            =   "5"
      Top             =   7710
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   6460
      Locked          =   -1  'True
      TabIndex        =   215
      Text            =   "5"
      Top             =   6390
      Width           =   60
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   214
      Text            =   "5"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   6435
      Locked          =   -1  'True
      TabIndex        =   213
      Text            =   "5"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   212
      Text            =   "5"
      Top             =   5090
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   24
      Left            =   5690
      Locked          =   -1  'True
      TabIndex        =   211
      Text            =   "4"
      Top             =   7230
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   210
      Text            =   "4"
      Top             =   7700
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   209
      Text            =   "4"
      Top             =   7720
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   21
      Left            =   3890
      Locked          =   -1  'True
      TabIndex        =   208
      Text            =   "4"
      Top             =   7400
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   207
      Text            =   "4"
      Top             =   7330
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   2220
      Locked          =   -1  'True
      TabIndex        =   206
      Text            =   "4"
      Top             =   7330
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   205
      Text            =   "4"
      Top             =   7330
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   690
      Locked          =   -1  'True
      TabIndex        =   204
      Text            =   "4"
      Top             =   7560
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   6590
      Locked          =   -1  'True
      TabIndex        =   203
      Text            =   "4"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   15
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   202
      Text            =   "4"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   3820
      Locked          =   -1  'True
      TabIndex        =   201
      Text            =   "4"
      Top             =   6360
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   200
      Text            =   "4"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   199
      Text            =   "4"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   198
      Text            =   "4"
      Top             =   5090
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   197
      Text            =   "4"
      Top             =   4830
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   3195
      Locked          =   -1  'True
      TabIndex        =   196
      Text            =   "4"
      Top             =   4350
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   2235
      Locked          =   -1  'True
      TabIndex        =   195
      Text            =   "4"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   194
      Text            =   "4"
      Top             =   4680
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   560
      Locked          =   -1  'True
      TabIndex        =   193
      Text            =   "4"
      Top             =   4940
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   192
      Text            =   "4"
      Top             =   3750
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   5720
      Locked          =   -1  'True
      TabIndex        =   191
      Text            =   "4"
      Top             =   3480
      Width           =   135
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   190
      Text            =   "4"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   189
      Text            =   "4"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   3210
      Locked          =   -1  'True
      TabIndex        =   188
      Text            =   "4"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   41
      Left            =   5620
      Locked          =   -1  'True
      TabIndex        =   187
      Text            =   "3"
      Top             =   840
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   40
      Left            =   5690
      Locked          =   -1  'True
      TabIndex        =   186
      Text            =   "3"
      Top             =   7010
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   39
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   185
      Text            =   "3"
      Top             =   7470
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   38
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   184
      Text            =   "3"
      Top             =   7480
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   37
      Left            =   3550
      Locked          =   -1  'True
      TabIndex        =   183
      Text            =   "3"
      Top             =   7600
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   36
      Left            =   2810
      Locked          =   -1  'True
      TabIndex        =   182
      Text            =   "3"
      Top             =   7710
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   35
      Left            =   2220
      Locked          =   -1  'True
      TabIndex        =   181
      Text            =   "3"
      Top             =   7000
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   180
      Text            =   "3"
      Top             =   7590
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   33
      Left            =   460
      Locked          =   -1  'True
      TabIndex        =   179
      Text            =   "3"
      Top             =   7560
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   32
      Left            =   6340
      Locked          =   -1  'True
      TabIndex        =   178
      Text            =   "3"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   31
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   177
      Text            =   "3"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   30
      Left            =   4640
      Locked          =   -1  'True
      TabIndex        =   176
      Text            =   "3"
      Top             =   6240
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   29
      Left            =   3820
      Locked          =   -1  'True
      TabIndex        =   175
      Text            =   "3"
      Top             =   6000
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   28
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   174
      Text            =   "3"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   173
      Text            =   "3"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   26
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   172
      Text            =   "3"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   25
      Left            =   630
      Locked          =   -1  'True
      TabIndex        =   171
      Text            =   "3"
      Top             =   5790
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   24
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   170
      Text            =   "3"
      Top             =   4850
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   169
      Text            =   "3"
      Top             =   4590
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   168
      Text            =   "3"
      Top             =   5080
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   21
      Left            =   3940
      Locked          =   -1  'True
      TabIndex        =   167
      Text            =   "3"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   2860
      Locked          =   -1  'True
      TabIndex        =   166
      Text            =   "3"
      Top             =   5070
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   2235
      Locked          =   -1  'True
      TabIndex        =   165
      Text            =   "3"
      Top             =   4350
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   164
      Text            =   "3"
      Top             =   4920
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   700
      Locked          =   -1  'True
      TabIndex        =   163
      Text            =   "3"
      Top             =   4560
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   6580
      Locked          =   -1  'True
      TabIndex        =   162
      Text            =   "3"
      Top             =   3360
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   4660
      Locked          =   -1  'True
      TabIndex        =   161
      Text            =   "3"
      Top             =   3750
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   160
      Text            =   "3"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   47
      Left            =   3205
      Locked          =   -1  'True
      TabIndex        =   159
      Text            =   "2"
      Top             =   3300
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   3820
      Locked          =   -1  'True
      TabIndex        =   158
      Text            =   "3"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   157
      Text            =   "3"
      Top             =   3750
      Width           =   135
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   2140
      Locked          =   -1  'True
      TabIndex        =   156
      Text            =   "3"
      Top             =   3750
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   155
      Text            =   "3"
      Top             =   3640
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   154
      Text            =   "3"
      Top             =   2120
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   46
      Left            =   5740
      Locked          =   -1  'True
      TabIndex        =   153
      Text            =   "2"
      Top             =   1830
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   5740
      Locked          =   -1  'True
      TabIndex        =   152
      Text            =   "3"
      Top             =   2280
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   4885
      Locked          =   -1  'True
      TabIndex        =   151
      Text            =   "3"
      Top             =   2040
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   5
      Left            =   3915
      Locked          =   -1  'True
      TabIndex        =   150
      Text            =   "3"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   3195
      Locked          =   -1  'True
      TabIndex        =   149
      Text            =   "3"
      Top             =   1700
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2260
      Locked          =   -1  'True
      TabIndex        =   148
      Text            =   "3"
      Top             =   1700
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   1540
      Locked          =   -1  'True
      TabIndex        =   147
      Text            =   "3"
      Top             =   2100
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   555
      Locked          =   -1  'True
      TabIndex        =   146
      Text            =   "3"
      Top             =   2320
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   45
      Left            =   6060
      Locked          =   -1  'True
      TabIndex        =   145
      Text            =   "2"
      Top             =   7230
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   44
      Left            =   5340
      Locked          =   -1  'True
      TabIndex        =   144
      Text            =   "2"
      Top             =   7600
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   43
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   143
      Text            =   "2"
      Top             =   7110
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   42
      Left            =   3560
      Locked          =   -1  'True
      TabIndex        =   142
      Text            =   "2"
      Top             =   7250
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   41
      Left            =   2810
      Locked          =   -1  'True
      TabIndex        =   141
      Text            =   "2"
      Top             =   7330
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   40
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   140
      Text            =   "2"
      Top             =   7710
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   39
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   139
      Text            =   "2"
      Top             =   7130
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   38
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   138
      Text            =   "2"
      Top             =   7080
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   37
      Left            =   6080
      Locked          =   -1  'True
      TabIndex        =   137
      Text            =   "2"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   36
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   136
      Text            =   "2"
      Top             =   5670
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   35
      Left            =   4640
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "2"
      Top             =   5840
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   3820
      Locked          =   -1  'True
      TabIndex        =   134
      Text            =   "2"
      Top             =   5680
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   33
      Left            =   3110
      Locked          =   -1  'True
      TabIndex        =   133
      Text            =   "2"
      Top             =   5800
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   32
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   132
      Text            =   "2"
      Top             =   6270
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   31
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "2"
      Top             =   5920
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   30
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "2"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   29
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   129
      Text            =   "2"
      Top             =   4620
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   28
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "2"
      Top             =   4350
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   127
      Text            =   "2"
      Top             =   4650
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   26
      Left            =   3580
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "2"
      Top             =   4920
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   25
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   125
      Text            =   "2"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   24
      Left            =   1875
      Locked          =   -1  'True
      TabIndex        =   124
      Text            =   "2"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   1060
      Locked          =   -1  'True
      TabIndex        =   123
      Text            =   "2"
      Top             =   4920
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   460
      Locked          =   -1  'True
      TabIndex        =   122
      Text            =   "2"
      Top             =   4560
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   21
      Left            =   6315
      Locked          =   -1  'True
      TabIndex        =   121
      Text            =   "2"
      Top             =   3360
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   5240
      Locked          =   -1  'True
      TabIndex        =   120
      Text            =   "2"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   4660
      Locked          =   -1  'True
      TabIndex        =   119
      Text            =   "2"
      Top             =   3480
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   3810
      Locked          =   -1  'True
      TabIndex        =   118
      Text            =   "2"
      Top             =   3030
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   2235
      Locked          =   -1  'True
      TabIndex        =   117
      Text            =   "2"
      Top             =   3300
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   16
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   116
      Text            =   "2"
      Top             =   3280
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   6110
      Locked          =   -1  'True
      TabIndex        =   115
      Text            =   "2"
      Top             =   2270
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   13
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   114
      Text            =   "2"
      Top             =   2430
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   12
      Left            =   3915
      Locked          =   -1  'True
      TabIndex        =   113
      Text            =   "2"
      Top             =   1700
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   11
      Left            =   2835
      Locked          =   -1  'True
      TabIndex        =   112
      Text            =   "2"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   1900
      Locked          =   -1  'True
      TabIndex        =   111
      Text            =   "2"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   9
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   110
      Text            =   "2"
      Top             =   2320
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   555
      Locked          =   -1  'True
      TabIndex        =   109
      Text            =   "2"
      Top             =   1880
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   6160
      Locked          =   -1  'True
      TabIndex        =   108
      Text            =   "2"
      Top             =   840
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   5260
      Locked          =   -1  'True
      TabIndex        =   107
      Text            =   "2"
      Top             =   840
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   106
      Text            =   "2"
      Top             =   1080
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   3920
      Locked          =   -1  'True
      TabIndex        =   105
      Text            =   "2"
      Top             =   820
      Width           =   135
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "2"
      Top             =   720
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   2380
      Locked          =   -1  'True
      TabIndex        =   103
      Text            =   "2"
      Top             =   720
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   1310
      Locked          =   -1  'True
      TabIndex        =   102
      Text            =   "2"
      Top             =   960
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   47
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   101
      Text            =   "1"
      Top             =   6990
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   46
      Left            =   5370
      Locked          =   -1  'True
      TabIndex        =   100
      Text            =   "1"
      Top             =   7100
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   45
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   99
      Text            =   "1"
      Top             =   7320
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   44
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   98
      Text            =   "1"
      Top             =   7000
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   43
      Left            =   3580
      Locked          =   -1  'True
      TabIndex        =   97
      Text            =   "1"
      Top             =   6990
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   42
      Left            =   1880
      Locked          =   -1  'True
      TabIndex        =   96
      Text            =   "1"
      Top             =   7330
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   41
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   95
      Text            =   "1"
      Top             =   7130
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   40
      Left            =   210
      Locked          =   -1  'True
      TabIndex        =   94
      Text            =   "1"
      Top             =   7330
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   39
      Left            =   6320
      Locked          =   -1  'True
      TabIndex        =   93
      Text            =   "1"
      Top             =   5670
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   38
      Left            =   5260
      Locked          =   -1  'True
      TabIndex        =   92
      Text            =   "1"
      Top             =   6120
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   37
      Left            =   4360
      Locked          =   -1  'True
      TabIndex        =   91
      Text            =   "1"
      Top             =   6120
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   36
      Left            =   3520
      Locked          =   -1  'True
      TabIndex        =   90
      Text            =   "1"
      Top             =   6120
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   35
      Left            =   2790
      Locked          =   -1  'True
      TabIndex        =   89
      Text            =   "1"
      Top             =   5800
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   34
      Left            =   2120
      Locked          =   -1  'True
      TabIndex        =   88
      Text            =   "1"
      Top             =   5790
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   33
      Left            =   1220
      Locked          =   -1  'True
      TabIndex        =   87
      Text            =   "1"
      Top             =   5670
      Width           =   195
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   32
      Left            =   160
      Locked          =   -1  'True
      TabIndex        =   86
      Text            =   "1"
      Top             =   6120
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   31
      Left            =   6100
      Locked          =   -1  'True
      TabIndex        =   85
      Text            =   "1"
      Top             =   4380
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   30
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   "1"
      Top             =   4800
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   29
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   83
      Text            =   "1"
      Top             =   4800
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   28
      Left            =   3580
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "1"
      Top             =   4350
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   27
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   81
      Text            =   "1"
      Top             =   4350
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   26
      Left            =   1900
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "1"
      Top             =   4350
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   25
      Left            =   1160
      Locked          =   -1  'True
      TabIndex        =   79
      Text            =   "1"
      Top             =   4520
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   24
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   78
      Text            =   "1"
      Top             =   4800
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   23
      Left            =   6080
      Locked          =   -1  'True
      TabIndex        =   77
      Text            =   "1"
      Top             =   3360
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   22
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   76
      Text            =   "1"
      Top             =   3030
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   21
      Left            =   4380
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "1"
      Top             =   3480
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   20
      Left            =   3560
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "1"
      Top             =   3480
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   19
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "1"
      Top             =   3300
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   18
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "1"
      Top             =   3300
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   17
      Left            =   1300
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   "1"
      Top             =   3020
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   15
      Left            =   6110
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "1"
      Top             =   1860
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   14
      Left            =   5350
      Locked          =   -1  'True
      TabIndex        =   69
      Text            =   "1"
      Top             =   2120
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   13
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   68
      Text            =   "1"
      Top             =   2040
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   12
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   67
      Text            =   "1"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   11
      Left            =   2860
      Locked          =   -1  'True
      TabIndex        =   66
      Text            =   "1"
      Top             =   1700
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   10
      Left            =   1900
      Locked          =   -1  'True
      TabIndex        =   65
      Text            =   "1"
      Top             =   1700
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Index           =   9
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   64
      Text            =   "1"
      Top             =   1920
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   8
      Left            =   230
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "1"
      Top             =   2100
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   7
      Left            =   6350
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "1"
      Top             =   380
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   6
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "1"
      Top             =   390
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   5
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "1"
      Top             =   480
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   3590
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "1"
      Top             =   820
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "1"
      Top             =   720
      Width           =   60
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      HideSelection   =   0   'False
      Index           =   2
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "1"
      Top             =   720
      Width           =   75
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   1310
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "1"
      Top             =   600
      Width           =   75
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   5690
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "6"
      Top             =   7700
      Width           =   135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   3195
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "5"
      Top             =   4800
      Width           =   60
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   2260
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "4"
      Top             =   2160
      Width           =   60
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   6565
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "3"
      Top             =   840
      Width           =   60
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "2"
      Top             =   800
      Width           =   135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "1"
      Top             =   800
      Width           =   75
   End
   Begin VB.CommandButton Cmd5 
      Caption         =   "Select"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Cmd2 
      Caption         =   "Define"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Color"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   6855
      Begin VB.OptionButton Radio18 
         Height          =   195
         Left            =   1200
         TabIndex        =   230
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio19 
         Height          =   195
         Left            =   2040
         TabIndex        =   229
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio20 
         Height          =   195
         Left            =   2880
         TabIndex        =   228
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio1 
         Height          =   195
         Left            =   360
         TabIndex        =   49
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio2 
         Height          =   195
         Left            =   1200
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio3 
         Height          =   195
         Left            =   2040
         TabIndex        =   47
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio4 
         Height          =   195
         Left            =   2880
         TabIndex        =   46
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio5 
         Height          =   195
         Left            =   3720
         TabIndex        =   45
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio6 
         Height          =   195
         Left            =   4560
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio7 
         Height          =   195
         Left            =   5400
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio8 
         Height          =   195
         Left            =   6240
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Radio41 
         Height          =   195
         Left            =   360
         TabIndex        =   40
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio48 
         Height          =   195
         Left            =   6240
         TabIndex        =   39
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio47 
         Height          =   195
         Left            =   5400
         TabIndex        =   38
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio46 
         Height          =   195
         Left            =   4560
         TabIndex        =   37
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio45 
         Height          =   195
         Left            =   3720
         TabIndex        =   36
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio44 
         Height          =   195
         Left            =   2880
         TabIndex        =   35
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio43 
         Height          =   195
         Left            =   2040
         TabIndex        =   34
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio42 
         Height          =   195
         Left            =   1200
         TabIndex        =   33
         Top             =   6840
         Width           =   255
      End
      Begin VB.OptionButton Radio33 
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio40 
         Height          =   195
         Left            =   6240
         TabIndex        =   31
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio39 
         Height          =   195
         Left            =   5400
         TabIndex        =   30
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio38 
         Height          =   195
         Left            =   4560
         TabIndex        =   29
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio37 
         Height          =   195
         Left            =   3720
         TabIndex        =   28
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio36 
         Height          =   195
         Left            =   2880
         TabIndex        =   27
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio35 
         Height          =   195
         Left            =   2040
         TabIndex        =   26
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio34 
         Height          =   195
         Left            =   1200
         TabIndex        =   25
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Radio25 
         Height          =   195
         Left            =   360
         TabIndex        =   24
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio32 
         Height          =   195
         Left            =   6240
         TabIndex        =   23
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio31 
         Height          =   195
         Left            =   5400
         TabIndex        =   22
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio30 
         Height          =   195
         Left            =   4560
         TabIndex        =   21
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio29 
         Height          =   195
         Left            =   3720
         TabIndex        =   20
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio28 
         Height          =   195
         Left            =   2880
         TabIndex        =   19
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio27 
         Height          =   195
         Left            =   2040
         TabIndex        =   18
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio26 
         Height          =   195
         Left            =   1200
         TabIndex        =   17
         Top             =   4200
         Width           =   255
      End
      Begin VB.OptionButton Radio17 
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio24 
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio23 
         Height          =   195
         Left            =   5400
         TabIndex        =   14
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio22 
         Height          =   195
         Left            =   4560
         TabIndex        =   13
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio21 
         Height          =   195
         Left            =   3720
         TabIndex        =   12
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Radio9 
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio16 
         Height          =   195
         Left            =   6240
         TabIndex        =   10
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio15 
         Height          =   195
         Left            =   5400
         TabIndex        =   9
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio14 
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio13 
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio12 
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio11 
         Height          =   195
         Left            =   2040
         TabIndex        =   5
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Radio10 
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   255
      End
      Begin VB.Line Line22 
         X1              =   4320
         X2              =   4800
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line120 
         X1              =   6000
         X2              =   6240
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line119 
         X1              =   6000
         X2              =   6240
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line118 
         X1              =   6000
         X2              =   6720
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line117 
         X1              =   6240
         X2              =   6240
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line116 
         X1              =   5640
         X2              =   5880
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line115 
         X1              =   5640
         X2              =   5880
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line114 
         X1              =   5640
         X2              =   5640
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line113 
         X1              =   5160
         X2              =   5880
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line112 
         X1              =   4800
         X2              =   5040
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line111 
         X1              =   4800
         X2              =   5040
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line110 
         X1              =   4800
         X2              =   4800
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line109 
         X1              =   3480
         X2              =   3720
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line108 
         X1              =   3480
         X2              =   3720
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line107 
         X1              =   3720
         X2              =   3720
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line106 
         X1              =   2640
         X2              =   3120
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line105 
         X1              =   3120
         X2              =   3120
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line104 
         X1              =   2640
         X2              =   3360
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line103 
         X1              =   2040
         X2              =   2520
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Line Line102 
         X1              =   1800
         X2              =   2520
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Line Line101 
         X1              =   2040
         X2              =   2040
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line100 
         X1              =   1200
         X2              =   1200
         Y1              =   7080
         Y2              =   7560
      End
      Begin VB.Line Line99 
         X1              =   960
         X2              =   1440
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line98 
         X1              =   1440
         X2              =   1440
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line97 
         X1              =   600
         X2              =   600
         Y1              =   7560
         Y2              =   8040
      End
      Begin VB.Line Line96 
         X1              =   360
         X2              =   840
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line95 
         X1              =   360
         X2              =   360
         Y1              =   7080
         Y2              =   8040
      End
      Begin VB.Line Line94 
         X1              =   6510
         X2              =   6510
         Y1              =   6000
         Y2              =   6480
      End
      Begin VB.Line Line93 
         X1              =   6210
         X2              =   6720
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line92 
         X1              =   6200
         X2              =   6200
         Y1              =   6000
         Y2              =   6720
      End
      Begin VB.Line Line91 
         X1              =   6000
         X2              =   6720
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line90 
         X1              =   5400
         X2              =   5400
         Y1              =   6000
         Y2              =   6720
      End
      Begin VB.Line Line89 
         X1              =   5160
         X2              =   5660
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line88 
         X1              =   5640
         X2              =   5640
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line87 
         X1              =   4500
         X2              =   4860
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line86 
         X1              =   4850
         X2              =   4850
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line85 
         X1              =   4480
         X2              =   4480
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line84 
         X1              =   3670
         X2              =   4010
         Y1              =   6400
         Y2              =   6400
      End
      Begin VB.Line Line83 
         X1              =   3670
         X2              =   4010
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line82 
         X1              =   4000
         X2              =   4000
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line81 
         X1              =   3670
         X2              =   3670
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line80 
         X1              =   3000
         X2              =   3000
         Y1              =   5760
         Y2              =   6240
      End
      Begin VB.Line Line79 
         X1              =   2640
         X2              =   3360
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line78 
         X1              =   2160
         X2              =   2160
         Y1              =   6240
         Y2              =   6720
      End
      Begin VB.Line Line77 
         X1              =   1800
         X2              =   2520
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line76 
         X1              =   1320
         X2              =   1320
         Y1              =   6240
         Y2              =   6720
      End
      Begin VB.Line Line75 
         X1              =   960
         X2              =   1680
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line74 
         X1              =   960
         X2              =   1680
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line73 
         X1              =   480
         X2              =   840
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Line Line72 
         X1              =   480
         X2              =   480
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line71 
         X1              =   290
         X2              =   290
         Y1              =   5760
         Y2              =   6720
      End
      Begin VB.Line Line70 
         X1              =   6000
         X2              =   6240
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line69 
         X1              =   6000
         X2              =   6240
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line68 
         X1              =   6000
         X2              =   6240
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line67 
         X1              =   6240
         X2              =   6240
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line66 
         X1              =   5640
         X2              =   5880
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line65 
         X1              =   5640
         X2              =   5880
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line64 
         X1              =   5640
         X2              =   5880
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line63 
         X1              =   5640
         X2              =   5640
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line62 
         X1              =   4800
         X2              =   5040
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line61 
         X1              =   4800
         X2              =   4800
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line60 
         X1              =   3480
         X2              =   3720
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line59 
         X1              =   3720
         X2              =   3720
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line58 
         X1              =   2640
         X2              =   3120
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line57 
         X1              =   3120
         X2              =   3120
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line56 
         X1              =   2640
         X2              =   3360
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line55 
         X1              =   2040
         X2              =   2520
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line54 
         X1              =   2040
         X2              =   2040
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line53 
         X1              =   1800
         X2              =   2520
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Line Line52 
         X1              =   1200
         X2              =   1200
         Y1              =   4920
         Y2              =   5400
      End
      Begin VB.Line Line51 
         X1              =   960
         X2              =   1440
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line50 
         X1              =   1440
         X2              =   1440
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line49 
         X1              =   600
         X2              =   600
         Y1              =   4440
         Y2              =   4920
      End
      Begin VB.Line Line48 
         X1              =   360
         X2              =   840
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line47 
         X1              =   360
         X2              =   360
         Y1              =   4440
         Y2              =   5400
      End
      Begin VB.Line Line46 
         X1              =   6520
         X2              =   6520
         Y1              =   3120
         Y2              =   3840
      End
      Begin VB.Line Line45 
         X1              =   6200
         X2              =   6200
         Y1              =   3120
         Y2              =   3840
      End
      Begin VB.Line Line44 
         X1              =   6000
         X2              =   6720
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line43 
         X1              =   5700
         X2              =   5700
         Y1              =   3360
         Y2              =   4080
      End
      Begin VB.Line Line42 
         X1              =   5350
         X2              =   5350
         Y1              =   3360
         Y2              =   4080
      End
      Begin VB.Line Line41 
         X1              =   5160
         X2              =   5880
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line40 
         X1              =   4500
         X2              =   4880
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line39 
         X1              =   4860
         X2              =   4860
         Y1              =   3120
         Y2              =   4080
      End
      Begin VB.Line Line38 
         X1              =   4500
         X2              =   4500
         Y1              =   3120
         Y2              =   4080
      End
      Begin VB.Line Line37 
         X1              =   3690
         X2              =   4020
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line36 
         X1              =   4000
         X2              =   4000
         Y1              =   3120
         Y2              =   4080
      End
      Begin VB.Line Line35 
         X1              =   3680
         X2              =   3680
         Y1              =   3120
         Y2              =   4080
      End
      Begin VB.Line Line34 
         X1              =   3120
         X2              =   3120
         Y1              =   3120
         Y2              =   3840
      End
      Begin VB.Line Line33 
         X1              =   2640
         X2              =   3360
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line32 
         X1              =   2040
         X2              =   2040
         Y1              =   3120
         Y2              =   3840
      End
      Begin VB.Line Line31 
         X1              =   1800
         X2              =   2520
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line30 
         X1              =   960
         X2              =   1680
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line29 
         X1              =   960
         X2              =   1680
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line28 
         X1              =   500
         X2              =   500
         Y1              =   3105
         Y2              =   4080
      End
      Begin VB.Line Line27 
         X1              =   300
         X2              =   300
         Y1              =   3120
         Y2              =   4080
      End
      Begin VB.Line Line26 
         X1              =   6000
         X2              =   6240
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line25 
         X1              =   5640
         X2              =   5880
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line24 
         X1              =   6240
         X2              =   6240
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line23 
         X1              =   5640
         X2              =   5640
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line21 
         X1              =   4800
         X2              =   4800
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line20 
         X1              =   3720
         X2              =   4200
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line19 
         X1              =   3720
         X2              =   3720
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line18 
         X1              =   3120
         X2              =   3120
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line17 
         X1              =   2640
         X2              =   3360
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line16 
         X1              =   2040
         X2              =   2040
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line15 
         X1              =   1800
         X2              =   2520
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line14 
         X1              =   960
         X2              =   1440
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line13 
         X1              =   1440
         X2              =   1440
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line12 
         X1              =   360
         X2              =   840
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line11 
         X1              =   360
         X2              =   360
         Y1              =   1800
         Y2              =   2760
      End
      Begin VB.Line Line10 
         X1              =   6480
         X2              =   6480
         Y1              =   720
         Y2              =   1440
      End
      Begin VB.Line Line9 
         X1              =   6000
         X2              =   6720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line8 
         X1              =   5400
         X2              =   5400
         Y1              =   720
         Y2              =   1440
      End
      Begin VB.Line Line7 
         X1              =   5160
         X2              =   5880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line6 
         X1              =   5040
         X2              =   4320
         Y1              =   1150
         Y2              =   1150
      End
      Begin VB.Line Line5 
         X1              =   3720
         X2              =   3720
         Y1              =   480
         Y2              =   1440
      End
      Begin VB.Line Line4 
         X1              =   2880
         X2              =   2880
         Y1              =   1440
         Y2              =   480
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   2280
         Y1              =   1440
         Y2              =   480
      End
      Begin VB.Line Line2 
         X1              =   960
         X2              =   1680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         X1              =   480
         X2              =   480
         Y1              =   480
         Y2              =   1440
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape4 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape5 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape6 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape7 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape8 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   480
         Width           =   735
      End
      Begin VB.Shape Shape48 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape47 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape46 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape45 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape44 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape43 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape42 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape41 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   7080
         Width           =   735
      End
      Begin VB.Shape Shape40 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape39 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape38 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape37 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape36 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape35 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape34 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape33 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   5760
         Width           =   735
      End
      Begin VB.Shape Shape32 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape31 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape30 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape29 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape28 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape27 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape26 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape25 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   4440
         Width           =   735
      End
      Begin VB.Shape Shape24 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape23 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape22 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape21 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape20 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape19 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape18 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape17 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   3120
         Width           =   735
      End
      Begin VB.Shape Shape16 
         BorderWidth     =   2
         Height          =   975
         Left            =   6000
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape15 
         BorderWidth     =   2
         Height          =   975
         Left            =   5160
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape13 
         BorderWidth     =   2
         Height          =   975
         Left            =   3480
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape12 
         BorderWidth     =   2
         Height          =   975
         Left            =   2640
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape11 
         BorderWidth     =   2
         Height          =   975
         Left            =   1800
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape10 
         BorderWidth     =   2
         Height          =   975
         Left            =   960
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape9 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   1800
         Width           =   735
      End
      Begin VB.Shape Shape14 
         BorderWidth     =   2
         Height          =   975
         Left            =   4320
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Choose The Frame Page You Wish To Make And Click The Radio Button On Top. Then Press The Select Button."
      Height          =   1935
      Left            =   6960
      TabIndex        =   41
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Frame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Repeat = "<HTML><HEAD><TITLE>Untitled Document</TITLE>" & vbCrLf & "<META HTTP-EQUIV=""Content-Type"" content=""text/html; charset=iso-8859-1"">" & vbCrLf & "</HEAD>" & vbCrLf
Private Const Repeat2 = "<NOFRAMES><BODY BGCOLOR=""FFFFFF"" TEXT=""000000"" LINK=""0000FF"" ALINK=""EEEEEE"" VLINK=""FF0000"">" & vbCrLf & "<P>This is a  frame page, but your browser doesn't support them.</P>" & vbCrLf & "</BODY></NOFRAMES></HTML>"
Private Sub ShowFrameTxt()
Frame.Hide
FrameText.Show
End Sub

Private Sub Cmd3_Click()

End Sub

Private Sub Cmd1_Click()
Frame.Hide
Color.Show
End Sub

Private Sub Cmd2_Click()
Dialog.Show
Frame.Hide
End Sub

Private Sub Cmd4_Click()
Frame.Hide
frmAbout.Show
End Sub

Private Sub Cmd5_Click()
If Radio1 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""100%,100%"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio2 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""50%,50%"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""Filename1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""Filename2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio3 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""*"" COLS=""80%,20%"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""Filename1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio4 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""*"" COLS=""20%,80%"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio5 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio6 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio7 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio8 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio9 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio10 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" BORDER=""0"" FRAMESPACING=""0"" ROWS=""*"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" BORDER=""0"" FRAMESPACING=""0"" COLS=""*"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio11 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio12 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio13 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""20%,80%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio14 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio15 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio16 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET><FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio17 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET><FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio18 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio19 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " </FRAMESET>" & vbCrLf & "<FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio20 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio21 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""20%,80%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""NO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio22 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""80%,20%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio23 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""NO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""NO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""NO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""NO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio24 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""ONE"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""TWO"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""THREE"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""FOUR"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio25 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio26 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio27 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""75%,25%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & _
" <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE><FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio28 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""25%,75%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & _
" <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio29 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio30 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio31 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & _
" <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio32 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & _
" <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio33 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio34 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio35 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio36 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio37 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,60%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE><FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET><FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio38 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%""COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio39 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""25%,75%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio40 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""75%,25%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""75%,25%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & _
"</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio41 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" TARGET="""">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" TARGET="""">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" TARGET="""">" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" TARGET="""">" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio42 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET COLS=""50%,50%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio43 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""25%,75%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & _
"</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio44 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""80%,20%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""25%,75%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & _
"</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio45 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio46 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio47 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""80%,20%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""40%,60%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & _
"<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName6.html"" FRAME NAME=""six"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
ElseIf Radio48 Then
ShowFrameTxt
FrameText.Txt = Repeat & "<FRAMESET COLS=""20%,80%"" ROWS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & "<FRAMESET ROWS=""40%,60%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName1.html"" FRAME NAME=""one"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "<FRAME SRC=""FileName2.html"" FRAME NAME=""two"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""50%,50%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName3.html"" FRAME NAME=""three"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName4.html"" FRAME NAME=""four"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & _
"</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & "<FRAMESET ROWS=""20%,80%"" COLS=""*"" BORDER=""0"" FRAMESPACING=""0"" BORDERCOLOR=""C0C0C0"">" & vbCrLf & " <FRAME SRC=""FileName5.html"" FRAME NAME=""five"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & " <FRAME SRC=""FileName6.html"" FRAME NAME=""six"" SCROLLING=""AUTO"" NORESIZE>" & vbCrLf & "</FRAMESET>" & vbCrLf & "</FRAMESET>" & vbCrLf & Repeat2
Else
MyVar = MsgBox("Please select a fame to copy the frame text from.", 0)
If MyVar = 1 Then Frame.Show
End If
End Sub

Private Sub Cmd6_Click()
On Error GoTo Err_Handler
If Cmd6.Caption = "Show Tips" Then
SaveSetting App.EXEName, "Options", "Show Tips at Startup", 1
frmTip.Show 0, Me
Cmd6.Caption = "Hide Tips"
ElseIf Cmd6.Caption = "Hide Tips" Then
Unload frmTip
SaveSetting App.EXEName, "Options", "Show Tips at Startup", 0
Unload frmTip
Cmd6.Caption = "Show Tips"
End If
Exit Sub
Err_Handler:

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then End
End Sub

Private Sub Form_Load()
Frame.Left = (Screen.Width / 2) - (Frame.Width / 2)
Frame.Top = (Screen.Height / 2) - (Frame.Height / 2)
Cmd6.Caption = "Show Tips"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
