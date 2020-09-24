VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About FragIt"
   ClientHeight    =   2820
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   3360
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1946.414
   ScaleMode       =   0  'User
   ScaleWidth      =   3155.214
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2535
      TabIndex        =   2
      Top             =   2325
      Width           =   645
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "For additional information email me at the following address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   4
      Left            =   165
      TabIndex        =   6
      Top             =   1920
      Width           =   2955
   End
   Begin VB.Label Label1 
      Height          =   150
      Left            =   2220
      TabIndex        =   5
      Top             =   180
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   84.515
      X2              =   3070.699
      Y1              =   496.957
      Y2              =   507.31
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lifeforcez@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   4
      Top             =   2310
      Width           =   1665
   End
   Begin VB.Label Label 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   165
      TabIndex        =   3
      Top             =   825
      Width           =   2955
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2002 by Muhammad Zubaer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   855
      TabIndex        =   1
      Top             =   495
      Width           =   2190
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FragIt"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   585
      TabIndex        =   0
      Top             =   210
      Width           =   990
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmAbout.frx":00B5
      Top             =   135
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

