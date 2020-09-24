VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Option Explicit

'==========Example=====================
'  Dim FileDialog As CFileDialog
'  Set FileDialog = New CFileDialog
'
'  With FileDialog
'    .DialogTitle = "Open"
'    .Filter = "All Files (*.*)|*.*"
'    .FilterIndex = 0
'    .Flags = FleFileMustExist + FleHideReadOnly + FleCreatePrompt
'    .hWndParent = Me.hwnd
'    If .Show(True) Then
'    MsgBox .FileTitle & " Was selected", , "Open"
'    Else
'    MsgBox "No file was selected", vbExclamation
'    Exit Sub
'    End If
'
'  End With
'===============End===========

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function GetOpenFileName _
  Lib "comdlg32.dll" _
  Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) _
  As Long

Private Declare Function GetSaveFileName _
  Lib "comdlg32.dll" _
  Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) _
  As Long

Private m_strDefaultExt As String
Private m_strDialogTitle As String
Private m_strFileName As String
Private m_strFileTitle As String
Private m_strInitialDir As String
Private m_strFilter As String
Private m_intFilterIndex As Integer
Private m_eFlags As EnumFilFlags
Private m_intMaxFileSize As Integer
Private m_lnghWndParent As Long

Private Const cintMaxFileLength As Integer = 260

Public Enum EnumFilFlags
  fleReadOnly = &H1
  fleOverWritePrompt = &H2
  FleHideReadOnly = &H4
  fleNoChangeDir = &H8
  fleShowHelp = &H10
  fleEnableHook = &H20
  fleEnableTemplate = &H40
  fleEnableTemplateHandle = &H80
  fleNoValidate = &H100
  fleAllowMultiSelect = &H200
  fleExtensionDifferent = &H400
  flePathMustExist = &H800
  FleFileMustExist = &H1000
  FleCreatePrompt = &H2000
  fleShareAware = &H4000
  fleNoReadOnlyReturn = &H8000
  fleNoTestFileCreate = &H10000
  fleNoNetworkButton = &H20000
  fleExplorer = &H80000
  fleLongnames = &H200000
End Enum

Public Property Get DefaultExt() As String
  DefaultExt = m_strDefaultExt

End Property

Public Property Let DefaultExt(ByVal strValue As String)
  m_strDefaultExt = strValue

End Property

Public Property Get DialogTitle() As String
  DialogTitle = m_strDialogTitle

End Property

Public Property Let DialogTitle(ByVal strValue As String)
  m_strDialogTitle = strValue

End Property

Public Property Get FileName() As String
  FileName = m_strFileName

End Property

Public Property Let FileName(ByVal strValue As String)

  m_strFileName = strValue

End Property

Public Property Get FileTitle() As String
  ' Returns: The filename without the path
  ' Source: Total VB SourceBook 6

  FileTitle = m_strFileTitle

End Property

Public Property Let FileTitle(ByVal strValue As String)
  m_strFileTitle = strValue

End Property

Public Property Get Filter() As String
  Filter = m_strFilter

End Property

Public Property Let Filter(ByVal strValue As String)
  m_strFilter = strValue

End Property

Public Property Get FilterIndex() As Integer
  FilterIndex = m_intFilterIndex

End Property

Public Property Let FilterIndex(ByVal intValue As Integer)
  m_intFilterIndex = intValue

End Property

Public Property Get Flags() As EnumFilFlags
  Flags = m_eFlags

End Property

Public Property Let Flags(ByVal eValue As EnumFilFlags)
  m_eFlags = eValue

End Property

Public Property Get hWndParent() As Long
  hWndParent = m_lnghWndParent

End Property

Public Property Let hWndParent(ByVal lngValue As Long)
  m_lnghWndParent = lngValue

End Property

Public Property Get InitialDir() As String
  InitialDir = m_strInitialDir

End Property

Public Property Let InitialDir(ByVal strValue As String)
  m_strInitialDir = strValue

End Property

Public Property Get MaxFileSize() As Integer
  MaxFileSize = m_intMaxFileSize

End Property

Public Property Let MaxFileSize(ByVal intValue As Integer)
  m_intMaxFileSize = intValue

End Property

Public Function Show(fOpen As Boolean) As Boolean
  Dim of As OPENFILENAME
  Dim strChar As String * 1
  Dim intCounter As Integer
  Dim strTemp As String
  
  On Error GoTo PROC_ERR
  
  ' Initialize the OPENFILENAME type
  of.lpstrTitle = m_strDialogTitle & ""
  of.Flags = m_eFlags
  of.lpstrDefExt = m_strDefaultExt & ""
  of.lStructSize = LenB(of)
  of.lpstrFilter = m_strFilter & "||"
  of.nFilterIndex = m_intFilterIndex
  
  ' To make Windows-style filter, replace pipes with nulls
  For intCounter = 1 To Len(m_strFilter)
    strChar = Mid$(m_strFilter, intCounter, 1)
    If strChar = "|" Then
      strTemp = strTemp & vbNullChar
    Else
      strTemp = strTemp & strChar
    End If
  Next
  
  ' Put double null at end
  strTemp = strTemp & vbNullChar & vbNullChar
  of.lpstrFilter = strTemp
  
  ' Pad file and file title buffers to maximum path length
  strTemp = m_strFileName & String$(cintMaxFileLength - Len(m_strFileName), 0)
  of.lpstrFile = strTemp
  of.nMaxFile = cintMaxFileLength
  
  strTemp = m_strFileTitle & String$(cintMaxFileLength - Len(m_strFileTitle), 0)
  of.lpstrFileTitle = strTemp
  of.lpstrInitialDir = m_strInitialDir
  of.nMaxFileTitle = cintMaxFileLength
  of.hwndOwner = m_lnghWndParent
  
  ' If fOpen is true, show the Open file dialog, otherwise show the Save dialog
  If fOpen Then
    If GetOpenFileName(of) Then
      Show = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  Else
    If GetSaveFileName(of) Then
      Show = True
      ' Assign property variables to appropriate values
      m_strFileName = TrimNulls(of.lpstrFile)
      m_strFileTitle = TrimNulls(of.lpstrFileTitle)
    Else
      Show = False
    End If
  End If
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
  "Show"
  Resume PROC_EXIT

End Function

Private Function TrimNulls(ByVal strIn As String) As String
    Dim intPos As Integer
  
  On Error GoTo PROC_ERR
    
  intPos = InStr(strIn, vbNullChar)
  
  If intPos = 0 Then
    ' No nulls in the string, just return it as is
    TrimNulls = strIn
  Else
    If intPos = 1 Then
      ' If the null character is at the first position, the
      ' entire string is a null string, so return a zero-length string
      TrimNulls = ""
    Else
      ' Not at the first position, so return the contents up
      ' to the occurrence of the null character
      TrimNulls = Left$(strIn, intPos - 1)
    End If
  End If
    
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TrimNulls"
  Resume PROC_EXIT
    
End Function




