VERSION 5.00
Begin VB.Form frmProgress 
   Caption         =   "����� �����"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_Select 
      Caption         =   "�������"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2340
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "����"
      Height          =   360
      Left            =   4260
      TabIndex        =   1
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3570
      ItemData        =   "frmProgress.frx":0000
      Left            =   120
      List            =   "frmProgress.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExitCode As String
Dim SearchPath As String
Dim SearchMask As String
Dim StopSearch As Boolean
Dim SearchInProgress As Boolean

Public Function ShowProgress(ByVal strPrompt As String, ByVal strSearchPath As String, ByVal strSearchMask As String) As String
'
' ��������� �����, ������� ��������� � ������ � ������
' ""- Cancel, else - ��������� ���������

ExitCode = ""
frmProgress.Caption = "����� """ & strSearchMask & """"
frmProgress.List1.Clear
frmProgress.List1.AddItem strPrompt
SearchPath = strSearchPath
SearchMask = strSearchMask
cmd_Cancel.Caption = "����"
cmd_Select.Enabled = False
frmProgress.Show 1, frmMain
ShowProgress = ExitCode

End Function

Private Sub FindFilesRec(ByVal path As String)
'
' ����������� ����� �����

Dim objName As String
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

Dim FileName As String
Dim FileExt As String
Dim LnkPath As String

Dim DoEventsCycle As Long

On Error Resume Next

WFD.cAlternate = 0

Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
  Do While Cont
    If StopSearch Then Exit Sub
    objName = Left$(WFD.cFileName, InStr(WFD.cFileName, Chr$(0)) - 1)
    If objName <> "." And objName <> ".." Then
      If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
        FindFilesRec path & objName & "\"
      Else
        ' ��������� ������
        If UCase$(Left$(WFD.cFileName, Len(SearchMask))) = UCase$(SearchMask) Then
          frmProgress.List1.AddItem path & WFD.cFileName
        End If
      End If
    End If
    Cont = FindNextFile(hSearch, WFD)
    ' ���������� ���������� ��� � 10 ������
    DoEventsCycle = DoEventsCycle + 1
    If DoEventsCycle >= 10 Then
      DoEventsCycle = 0
      DoEvents
    End If
  Loop
  Cont = FindClose(hSearch)
End If

End Sub

Private Sub cmd_Cancel_Click()
'
' �������� / ���� - � ����������� �� ���� ������� ����� ��� ��� ����������

If SearchInProgress Then
  StopSearch = True
  cmd_Cancel.Caption = "������"
  cmd_Select.Enabled = True
Else
  ExitCode = ""
  StopSearch = True
  frmProgress.Hide
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'
' ���� ��� ����������, ��������� ����� ������

If UnloadMode = vbFormOwner Then
  Cancel = False
  Exit Sub
End If

If SearchInProgress Then
  StopSearch = True
  cmd_Cancel.Caption = "������"
  cmd_Select.Enabled = True
  Cancel = True
Else
  StopSearch = False
  If List1.ListCount > 0 Then List1.RemoveItem 0
  Cancel = False
End If

End Sub

Private Sub cmd_Select_Click()
'
' �������

Dim SelItem As String

If SearchInProgress Then Exit Sub

SelItem = List1.List(List1.ListIndex)
If (SelItem <> "") And (SelItem <> msgNothingFound) Then
  ExitCode = SelItem
End If
frmProgress.Hide

End Sub

Private Sub Form_Activate()
'
' ���� ��� ����������, ��������� ����� ������

StopSearch = False
SearchInProgress = True
' ������ ������
FindFilesRec SearchPath
SearchInProgress = False
cmd_Cancel.Caption = "������"
List1.RemoveItem 0
If List1.ListCount = 0 Then
  List1.AddItem msgNothingFound
Else
  cmd_Select.Enabled = True
End If

End Sub

Private Sub List1_DblClick()
'
' ������� ����

Call cmd_Select_Click

End Sub

Sub Form_Resize()
'
' ��������������� ����

If Me.WindowState <> 1 Then
  List1.Width = frmProgress.Width - 442
  List1.Height = frmProgress.Height - 1215
  cmd_Select.Top = frmProgress.Height - 1115
  cmd_Cancel.Top = frmProgress.Height - 1115
  cmd_Select.Left = (frmProgress.Width / 2) - cmd_Select.Width - 50
  cmd_Cancel.Left = (frmProgress.Width / 2) + 50
End If

End Sub
