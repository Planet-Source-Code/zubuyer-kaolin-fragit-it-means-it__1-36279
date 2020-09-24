VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "FragIt"
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   45
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   187
      TabIndex        =   7
      Top             =   1020
      Width           =   2835
   End
   Begin VB.PictureBox picCHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   0
      Left            =   45
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   0
      Top             =   270
      Width           =   3510
      Begin VB.TextBox txtSegSize 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         TabIndex        =   3
         Text            =   "1.38"
         ToolTipText     =   "Text field for segment file size"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtSFName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         ToolTipText     =   "Text field for file name"
         Top             =   45
         Width           =   2055
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MB"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3195
         TabIndex        =   12
         Top             =   405
         Width           =   240
      End
      Begin VB.Label lblBrowse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3165
         MouseIcon       =   "frmMain.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   5
         ToolTipText     =   "Browse"
         Top             =   45
         Width           =   270
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fragment Size"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   4
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   90
         Width           =   705
      End
   End
   Begin VB.PictureBox picCHolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   720
      Index           =   1
      Left            =   45
      ScaleHeight     =   690
      ScaleWidth      =   3480
      TabIndex        =   8
      Top             =   270
      Width           =   3510
      Begin VB.TextBox txtSegfname 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1065
         OLEDropMode     =   1  'Manual
         TabIndex        =   9
         ToolTipText     =   "Text field for segment file to be merged"
         Top             =   45
         Width           =   2055
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fragment file"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   90
         Width           =   900
      End
      Begin VB.Label lblSBrowse 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "..."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3165
         MouseIcon       =   "frmMain.frx":074C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         ToolTipText     =   "Browse"
         Top             =   45
         Width           =   270
      End
   End
   Begin VB.Image imgExit 
      Height          =   165
      Left            =   3375
      MouseIcon       =   "frmMain.frx":0A56
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0D60
      ToolTipText     =   "Exit fSplit"
      Top             =   60
      Width           =   180
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   165
      Left            =   2670
      Picture         =   "frmMain.frx":0E6E
      Top             =   75
      Width           =   465
   End
   Begin VB.Image imgInfo 
      Height          =   165
      Left            =   3180
      MouseIcon       =   "frmMain.frx":0FD4
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":12DE
      ToolTipText     =   "About fSplit"
      Top             =   60
      Width           =   180
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00000000&
      Height          =   1320
      Left            =   0
      Top             =   0
      Width           =   3600
   End
   Begin VB.Label lblInitialize 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Initialize"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2910
      MouseIcon       =   "frmMain.frx":13EC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      ToolTipText     =   "Initialize file split or merge"
      Top             =   1020
      Width           =   645
   End
   Begin VB.Image imgTab 
      Height          =   240
      Index           =   0
      Left            =   45
      MouseIcon       =   "frmMain.frx":16F6
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":1A00
      Top             =   45
      Width           =   720
   End
   Begin VB.Image imgTab 
      Height          =   240
      Index           =   1
      Left            =   675
      MouseIcon       =   "frmMain.frx":1C8A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":1F94
      Top             =   45
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''
'FragIt                                '
'© Copyright 2002 by Muhammad Zubaer   '
'                                      '
'This is a FREEWARE but this code      '
'is not intend to be used commercially.'
'Although you can use it as you like   '
'in your own project but do not resale '
'it or destroy the original author's   '
'name. If you use this code in your    '
'project than it would be nice to give '
'me some cradits. I've worked hard on  '
'it.
'                                      '
'Warning: There is no warranty provided'
'so use it in your own risk. The author'
'is not responsible for any damage     '
'caused by this code.                  '
'                                      '
'Mail me at the following address if   '
'you have any questions or made any    '
'enhancement.                          '
'lifeforcez@hotmail.com                '
''''''''''''''''''''''''''''''''''''''''

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim Mode As Integer

Private Sub Form_Load()
  Mode = 0
  spProg 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 1
    End If
End Sub

Private Sub imgInfo_Click()
frmAbout.Show 1, Me
End Sub

Private Sub imgTab_Click(Index As Integer)
Mode = Index
picCHolder(Index).ZOrder
imgTab(Index).ZOrder
End Sub

Private Sub lblBrowse_Click()
  Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  With FileDialog
    .DialogTitle = "Open"
    .Filter = "All Files (*.*)|*.*"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hWnd
    If .Show(True) Then
    txtSFName.Text = .FileName
    Else
    Exit Sub
    End If
  End With
End Sub

Private Sub lblBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.BackColor = 12632256
End Sub

Private Sub lblBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBrowse.BackColor = 14737632
End Sub

Private Sub imgExit_Click()
Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.BackColor = 12632256
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblExit.BackColor = 14737632
End Sub

Private Sub lblInitialize_Click()
Dim i, j As Integer
Dim Segments As Integer

lblInitialize.Enabled = False
Select Case Mode
Case 0
    
    'Call the function
    i = SplitFile(txtSFName.Text, val(txtSegSize.Text) * 1048576, Segments)
    'Inform the user about the call success or failure
    If i = 0 Then
        MsgBox "The process completed successfully." & Chr(10) & "The file was split to " & Segments & " segments.", vbExclamation, "Successfully Done"
    Else
        MsgBox "An error occured!" & vbCr & "Try entering different segment value or check if the file name is correct.", vbExclamation, "Error"
    End If
Case 1
    
    'Call the function
    j = MergeFiles(txtSegfname.Text, Segments)
 
    'Inform the user about the call success or failure
    If j = 0 Then
        MsgBox "The process completed successfully." & Chr(10) & "The file was merged from " & Segments & " segments.", vbExclamation, "Successfully Done"
    Else
        MsgBox "An error occured!" & Chr(10) & "Check if all the fraggment files are in the same directory" & Chr(10) & "or make sure the program is not overwriting any file.", vbExclamation, "Error"
    End If

End Select
lblInitialize.Enabled = True
End Sub

Private Sub lblInitialize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInitialize.BackColor = 12632256
End Sub

Private Sub lblInitialize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblInitialize.BackColor = 14737632
End Sub

Private Sub lblSBrowse_Click()
  Dim FileDialog As CFileDialog
  Set FileDialog = New CFileDialog
  With FileDialog
    .DialogTitle = "Open"
    .Filter = "All Files (*.*)|*.*"
    .FilterIndex = 0
    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
    .hWndParent = Me.hWnd
    If .Show(True) Then
    txtSegfname.Text = .FileName
    Else
    Exit Sub
    End If
  End With
End Sub

Public Sub spProg(val As Integer)
picProgress.Cls
picProgress.Line (30, 2)-((val * 1.55) + 30, 12), &H808080, BF
picProgress.Line (30, 2)-(184, 12), 0, B
picProgress.CurrentX = 3: picProgress.CurrentY = 1
picProgress.Print val & "%"
End Sub

Private Sub lblSBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSBrowse.BackColor = 12632256
End Sub

Private Sub lblSBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSBrowse.BackColor = 14737632
End Sub

Private Sub txtSegfname_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtSegfname.Text = Data.Files(1)
End Sub

Private Sub txtSFName_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtSFName.Text = Data.Files(1)
End Sub
