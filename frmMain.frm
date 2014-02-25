VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "AnalizeLNK - https://github.com/wyfinger/AnalizeLNK"
   ClientHeight    =   10170
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "��������"
      Height          =   400
      Left            =   15480
      TabIndex        =   5
      Top             =   9600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   9600
      Width           =   14655
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "�����"
      Height          =   400
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7980
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      MultiSelect     =   2  '����������
      TabIndex        =   1
      Top             =   1400
      Width           =   16455
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   2040
      TabIndex        =   0
      Text            =   "\\Prim-fs-serv\rdu\����\���\�����\"
      Top             =   100
      Width           =   14535
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  '�������
      Height          =   400
      Left            =   14880
      Top             =   9600
      Width           =   500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   20000
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Label Label1 
      Caption         =   "������� ��� ���������:"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "���������: 0 ������, 0 ������, �� ��� ����� 0"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   945
      Width           =   7095
   End
   Begin VB.Menu pm_List 
      Caption         =   "�"
      Visible         =   0   'False
      Begin VB.Menu mi_OpenLnkFile 
         Caption         =   "������� ������� � ������ ������"
      End
      Begin VB.Menu mi_OpenFarFolder 
         Caption         =   "������� ��������� ������������ �������"
      End
      Begin VB.Menu mi_FindDest 
         Caption         =   "����� <>"
      End
      Begin VB.Menu mi_LnkProperties 
         Caption         =   "�������� ����� ������"
      End
      Begin VB.Menu mi_SaveList 
         Caption         =   "��������� ������ ������� � ����"
         Begin VB.Menu mi_SaveLnkFilesList 
            Caption         =   "������ Lnk �����"
         End
         Begin VB.Menu mi_SaveDestFilesList 
            Caption         =   "������ ����� �������"
         End
         Begin VB.Menu mi_SaveBothList 
            Caption         =   "Lnk ��� <-> ����"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StopSearch As Boolean ' ���������� �����
Dim SearchState As Boolean ' True - ��������, False - �� ��������
Dim FilesCount As Long
Dim LnkCount As Long
Dim BadLnk As Long
Dim BothPart As String
Dim InitialDir As String

Private Function ExtractFileExt(ByVal strFileName As String) As String
'
' ��������� ���������� �����
 
Dim strUp As String
Dim dotPoint As Long
strUp = UCase$(strFileName) ' ��������� �������
dotPoint = InStrRev(strUp, ".")
If dotPoint > 0 Then
  ExtractFileExt = Right$(strUp, Len(strUp) - dotPoint)
Else
  ExtractFileExt = ""
End If
  
End Function

Private Function ExtractFilePath(ByVal strPath As String) As String
'
' ��������� �������� �����

Dim slash_pos As Long

slash_pos = InStrRev(Replace$(strPath, "/", "\"), "\")
If slash_pos > 0 Then
  ExtractFilePath = Left$(strPath, Len(strPath) - slash_pos)
Else
  ExtractFilePath = ""
End If
  
End Function

Private Function ExtractFileName(ByVal strPath As String) As String
'
' ��������� ����� �����

Dim slash_pos As Long

If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)

slash_pos = InStrRev(Replace$(strPath, "/", "\"), "\")
If slash_pos > 0 Then
  ExtractFileName = Right$(strPath, Len(strPath) - slash_pos)
Else
  ExtractFileName = ""
End If
  
End Function

Private Function FileExists(ByVal strFileName As String) As Boolean
'
' �������� ������������� �����

FileExists = PathFileExists(strFileName)
  
End Function

Private Function GetLinkPath(ByVal lnk As String) As String
'
' ��������� ����, �� ������� ��������� �����

GetLinkPath = ""
On Error Resume Next
  With CreateObject("Wscript.Shell").CreateShortcut(lnk)
    GetLinkPath = .TargetPath
    .Close
  End With
  
End Function

Private Function IsRealyLnk(ByVal LnkFile As String) As Boolean
'
' �������� ������������� �� ���� ���� �������� Shell Link (.LNK) Binary File
' ��. http://msdn.microsoft.com/en-us/library/dd871305.aspx

Dim SFile As Integer
Dim Readed4 As Long

SFile = FreeFile
Open LnkFile For Binary Access Read As SFile
Get SFile, , Readed4
' ������ 4 ����� ������ ���� 0x0000004C
If Readed4 <> 76 Then Exit Function
' ��������� 16 ���� ������ ���� 00021401-0000-0000-C000-000000000046
Get SFile, , Readed4: If Readed4 <> 136193 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 0 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 192 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 1174405120 Then Exit Function
Close SFile

IsRealyLnk = True

Exit Function
err:
  ' ��� ������� ��� ��� ����-������, �� �����, ��� ��� ��� �� �����
  IsRealyLnk = False
End Function

Private Sub ProcessFiles(ByVal path As String)
'
' ������� ����������� � ������� LNK ������

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
        ProcessFiles path & objName & "\"
      Else
        ' ��������� ������
        FilesCount = FilesCount + 1
        FileName = objName
        FileExt = ExtractFileExt(objName)
        If (FileExt = "LNK") Then
          If IsRealyLnk(path & FileName) Then  ' � VB6 ��� ���������� ����������!
            LnkCount = LnkCount + 1
            LnkPath = GetLinkPath(path & FileName)
            If Not FileExists(LnkPath) Then
              BadLnk = BadLnk + 1
              List1.AddItem (path & FileName)
            End If
          End If
        End If
      End If
    End If
    Cont = FindNextFile(hSearch, WFD)
    Label2.Caption = "���������� " & FilesCount & " ������, ������� " & LnkCount & ", �� ��� ����� " & BadLnk
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

Private Function SaveFileDialog(InitialDir As String) As String
'
' �������� ������ ������ ����� ��� �����������

Dim OFName As OPENFILENAME
If InitialDir = "" Then InitialDir = App.path

OFName.lStructSize = Len(OFName)
OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "��� ����� (*.*)" + Chr$(0) + "*.*" + Chr$(0)
OFName.lpstrFile = Space$(254)
OFName.nMaxFile = 255
OFName.lpstrFileTitle = Space$(254)
OFName.nMaxFileTitle = 255
OFName.lpstrInitialDir = InitialDir
OFName.lpstrTitle = "���������� �����"
OFName.flags = 0
 
If GetSaveFileName(OFName) Then
  SaveFileDialog = Trim$(OFName.lpstrFile)
Else
  SaveFileDialog = ""
End If

End Function

Private Sub SetLinkPath(ByVal LnkFile As String, ByVal lnk As String)
'
' ��������� ����, �� ������� ��������� �����

On Error Resume Next
  With CreateObject("Wscript.Shell").CreateShortcut(LnkFile)
    .TargetPath = lnk
    .Save
    .Close
  End With
  
End Sub

Sub Form_Resize()
'
' ��������������� ����

If Me.WindowState <> 1 Then
  Text1.Width = frmMain.Width - 2415
  List1.Width = frmMain.Width - 495
  Text2.Width = frmMain.Width - 1695 - 500
  Shape1.Left = frmMain.Width - 1695 - 350
  Command3.Left = frmMain.Width - 1470
  List1.Height = frmMain.Height - 2760
  Text2.Top = frmMain.Height - 1140
  Shape1.Top = frmMain.Height - 1140
  Command3.Top = frmMain.Height - 1140
  Line1.X2 = frmMain.Width + 5000
End If

End Sub

Sub mi_FindDest_Click()
'
' ����� ����� (�� �����) �� ������� ��������� ����� � ���������� ������

Dim LnkFile As String
Dim LnkDest As String
Dim DestName As String
Dim Selected As String

LnkFile = List1.List(List1.ListIndex)
LnkDest = GetLinkPath(LnkFile)
DestName = ExtractFileName(LnkDest)

If DestName <> "" Then
  ' �������� ����� ������� �����
  Selected = frmProgress.ShowProgress("���� �����, �����...", Text1.Text, DestName)
End If

If Selected <> "" Then
  MsgBox Selected
End If

End Sub

Sub mi_LnkProperties_Click()
'
' ������� ������ ������� ����� ������

Dim LnkFile As String
Dim SEI As SHELLEXECUTEINFO

LnkFile = List1.List(List1.ListIndex)
If Not FileExists(LnkFile) Then Exit Sub

With SEI
  .cbSize = Len(SEI)
  .fMask = SEE_MASK_INVOKEIDLIST
  .hwnd = Me.hwnd
  .lpVerb = "properties"
  .lpFile = LnkFile
  .lpParameters = vbNullChar
  .lpDirectory = vbNullChar
  .nShow = 0
  .hInstApp = 0
  .lpIDList = 0
End With

Call ShellExecuteEx(SEI)

End Sub

Sub mi_SaveBothList_Click()
'
' ���������� ������ LNK ������ � �� �����

Dim FileToSave As String
Dim SFile As Integer
Dim i As Long
SFile = FreeFile

On Error GoTo err
FileToSave = SaveFileDialog(InitialDir)
InitialDir = ExtractFilePath(FileToSave)
If FileToSave <> "" Then
  Open FileToSave For Output As SFile
  For i = 0 To List1.ListCount - 1
    Print #SFile, List1.List(i) & " <-> " & GetLinkPath(List1.List(i))
  Next
  Close SFile
End If

Exit Sub
err:
  MsgBox err.Description
End Sub

Sub mi_SaveDestFilesList_Click()
'
' ���������� ������ ����� �������

Dim FileToSave As String
Dim SFile As Integer
Dim i As Long
SFile = FreeFile

On Error GoTo err
FileToSave = SaveFileDialog(InitialDir)
InitialDir = ExtractFilePath(FileToSave)
If FileToSave <> "" Then
  Open FileToSave For Output As SFile
  For i = 0 To List1.ListCount - 1
    Print #SFile, GetLinkPath(List1.List(i))
  Next
  Close SFile
End If

Exit Sub
err:
  MsgBox err.Description
End Sub

Sub mi_SaveLnkFilesList_Click()
'
' ���������� ������ LNK ������

Dim FileToSave As String
Dim SFile As Integer
Dim i As Long
SFile = FreeFile

On Error GoTo err
FileToSave = SaveFileDialog(InitialDir)
InitialDir = ExtractFilePath(FileToSave)
If FileToSave <> "" Then
  Open FileToSave For Output As SFile
  For i = 0 To List1.ListCount - 1
    Print #SFile, List1.List(i)
  Next
  Close SFile
End If

Exit Sub
err:
  MsgBox err.Description
End Sub

Private Sub cmdStartStop_Click()
'
' ���������

If Right$(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"

If SearchState Then
  StopSearch = True
  cmdStartStop.Caption = "�����"
Else
  StopSearch = False
  FilesCount = 0
  LnkCount = 0
  BadLnk = 0
  List1.Clear
  cmdStartStop.Caption = "����"
  SearchState = True
  ProcessFiles Text1.Text
  SearchState = False
  cmdStartStop.Caption = "�����"
  Label2.Caption = Label2.Caption & ". ���������."
End If

End Sub

Private Sub Command3_Click()
'
' ������ ��������� � �����

Dim LnkFile As String
Dim lnk As String
Dim i As Long

' ���� ��� ������� ���� ���� - ������ ������� ��� ������,
' ���� �������� ��������� ������ - ������� � ������������ ��� ���������:
' 1. �������� � ���� ���������� ������� ����� ����� � ����
' 2. �������� � ���� ���������� ������� ��� ������ ������� (����� ���������
' ������� ��������� �� ���� ����

If List1.SelCount = 1 Then
  LnkFile = List1.List(List1.ListIndex)
  If FileExists(LnkFile) Then
    SetLinkPath LnkFile, Text2.Text
  End If
Else
  ' ���������� ������������� ��������� ��� ������������, ����� �� ���������
  Select Case frmQuery.QueryMode("� ��������, �� ������� ��������� ���������� ���� ������ ������� " & _
  "��������� ����� �����:" & vbCrLf & BothPart & vbCrLf & vbCrLf & _
  "�������� � ��������� ������� ����� ����� ���� �� �����, ��� ������ �������� ������?")
    Case 0
      ' ������ �� ������
    Case 1
      ' ������ ����� ����� � �����
      For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          LnkFile = List1.List(i)
          lnk = GetLinkPath(LnkFile)
          lnk = Text2.Text & Mid$(lnk, Len(BothPart) + 1, Len(lnk) - Len(BothPart))
          SetLinkPath LnkFile, lnk
        End If
      Next
    Case 2
      ' ������ ������ ����� �� �����
      For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          LnkFile = List1.List(i)
          SetLinkPath LnkFile, Text2.Text
        End If
      Next
  End Select
End If

End Sub

Private Sub List1_Click()
'
' ���� �� ������ ������ - �� ������ ������

Dim SelectedLinks As New Collection
Dim LnkFile As String
Dim LnkPath As String
Dim i As Long
Dim j As Long
Dim p As Long
Dim CharA As String
Dim CharB As String
Dim FindDiff As Boolean

' ���� ������� ���� ����� - ������ ��������� ������ �� ������
' � ��������� ������ ������� ����� ����� BothPart
If List1.SelCount = 1 Then
  BothPart = ""
  LnkFile = List1.List(List1.ListIndex)
  If FileExists(LnkFile) Then
    LnkPath = GetLinkPath(LnkFile)
    Text2.Text = LnkPath
  End If
Else
  For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
      SelectedLinks.Add GetLinkPath(List1.List(i))
    End If
  Next
  ' ���� ����� �����
  p = 1
  FindDiff = False
  Do While Not FindDiff
    ' ����� ��������� ������ �� ������ ������, � ������� � ����������
    CharA = Mid$(SelectedLinks(1), p, 1)
    For j = 2 To List1.SelCount
      CharB = Mid$(SelectedLinks(j), p, 1)
      If UCase$(CharA) <> UCase$(CharB) Then
        FindDiff = True
        Exit For
      End If
    Next
    p = p + 1
  Loop
  BothPart = Mid$(SelectedLinks(1), 1, p - 2)
  ' ������� ������� �� ������ "\" � ����� ������ �� ������, ���� ������ ����
  ' ������ ��� ��������� ���������
  BothPart = Left$(BothPart, InStrRev(BothPart, "\"))
  Text2.Text = BothPart
End If

Set SelectedLinks = Nothing

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'
' ������� Del �� ������, �������� ������ �������

Dim OneLnk As String
Dim LnkCount As Long
Dim MustDel As Long
Dim i As Long

If KeyCode = 46 Then
  ' ������� ���������� ���������� ������
  If List1.SelCount = 1 Then
    OneLnk = List1.List(List1.ListIndex)
    MustDel = MsgBox("������� ���� ������:" & vbCrLf & OneLnk, vbYesNo, "������� �����")
  Else
    LnkCount = List1.SelCount
    MustDel = MsgBox("�������� " & LnkCount & " ������, �������?", vbYesNo, "������� �����")
  End If
  If MustDel = 6 Then
    For i = List1.ListCount - 1 To 0 Step -1
      If List1.Selected(i) Then
        OneLnk = List1.List(i)
        If FileExists(OneLnk) Then
          DeleteFile (OneLnk)
          List1.RemoveItem i
        End If
      End If
    Next
  End If
End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
'
' ��� ������� ����� ������ ���������� ���������� ����

Call List1_Click

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' ���������� ���� �� ������ ������

If Button = 2 Then
  ' ������� ������� ��� ����� ������ ������� ����
  ' ������ ��� �����
  ' List1.TopIndex
  
  Dim LnkFile As String
  Dim LnkDest As String
 
  If (List1.SelCount = 1) And (FileExists(List1.List(List1.TabIndex))) Then
    mi_OpenLnkFile.Enabled = True
    mi_OpenFarFolder.Enabled = True
    LnkFile = List1.List(List1.ListIndex)
    If FileExists(LnkFile) Then
      LnkDest = GetLinkPath(LnkFile)
      If LnkDest <> "" Then
        mi_FindDest = True
        mi_FindDest.Caption = "����� """ & ExtractFileName(LnkDest) & """"
      End If
    End If
  Else
    mi_OpenLnkFile.Enabled = False
    mi_OpenFarFolder.Enabled = False
  End If
  If List1.ListCount > 0 Then
    mi_SaveList.Enabled = True
  Else
    mi_SaveList.Enabled = False
  End If
  PopupMenu pm_List, vbPopupMenuRightButton
End If

End Sub

Private Sub mi_OpenFarFolder_Click()
'
' ���������� ��������������� ����� ��������, ���� ��������� ����� � ��������� ��������� ������������

Dim LnkFile As String
Dim LnkPath As String
Dim SlashPos As Long

LnkFile = List1.List(List1.ListIndex)
If FileExists(LnkFile) Then
  LnkPath = GetLinkPath(LnkFile)
    
  SlashPos = InStrRev(LnkPath, "\")
  If SlashPos = 0 Then Exit Sub
  
  Do While (SlashPos > 0) And (Not FileExists(LnkPath))
    LnkPath = Left$(LnkPath, SlashPos - 1)
    SlashPos = InStrRev(LnkPath, "\")
    If SlashPos = 0 Then Exit Sub
  Loop
  
  If FileExists(LnkPath) Then ShellExecute frmMain.hwnd, "OPEN", "EXPLORER", LnkPath, 0, SW_SHOWNORMAL
End If
 
End Sub

Private Sub mi_OpenLnkFile_Click()
'
' ������� ������� � ������� � �������� ���

Dim LnkFile As String

LnkFile = List1.List(List1.ListIndex)
If FileExists(LnkFile) Then ShellExecute frmMain.hwnd, "OPEN", "EXPLORER", "/select, " & LnkFile, 0, SW_SHOWNORMAL
  
End Sub


Private Sub Text2_Change()
'
' ��� �������������� ����, ���� ��������� ����� ��������� ���������� �� ����
' ���� ��� �����

If FileExists(Text2.Text) Then
  Shape1.FillColor = &H8000&
Else
   Shape1.FillColor = &HC0&
End If

End Sub
