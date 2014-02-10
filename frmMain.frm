VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Проверка ярлыков"
   ClientHeight    =   10170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Ближайший существующий каталог"
      Height          =   400
      Left            =   13200
      TabIndex        =   8
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Каталог с ярлыком"
      Height          =   400
      Left            =   11040
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Изменить"
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
      Width           =   15255
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Старт"
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
      MultiSelect     =   2  'Расширенно
      TabIndex        =   1
      Top             =   1400
      Width           =   16455
   End
   Begin VB.TextBox Text1 
      Height          =   400
      Left            =   2040
      TabIndex        =   0
      Text            =   "\\Prim-fs-serv\rdu\СРЗА\ТКЗ\Линии\"
      Top             =   100
      Width           =   14535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000006&
      X1              =   0
      X2              =   20000
      Y1              =   650
      Y2              =   650
   End
   Begin VB.Label Label1 
      Caption         =   "Каталог для обработки:"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Обработка: 0 файлов, 0 ярлыка, из них битых 0"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   945
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal path As String) As Boolean

Const SW_SHOWNORMAL = 1

Dim StopSearch As Boolean ' остановить поиск
Dim SearchState As Boolean ' True - работаем, False - не работаем
Dim FilesCount As Long
Dim LnkCount As Long
Dim BadLnk As Long

Private Function GetLinkPath(ByVal lnk As String)
'
' Получение пути, на который ссылается ярлык

GetLinkPath = ""
On Error Resume Next
  With CreateObject("Wscript.Shell").CreateShortcut(lnk)
    GetLinkPath = .TargetPath
    .Close
  End With
  
End Function

Private Function FileExists(ByVal strFileName As String) As Boolean
'
' Проверка существования файла

FileExists = PathFileExists(strFileName)
  
End Function

Private Function ExtractFileExt(ByVal strFileName As String) As String
'
' Получение расширения файла
 
Dim strUp As String
Dim dotPoint As Long
strUp = UCase$(strFileName) ' поднимаем регистр
dotPoint = InStrRev(strUp, ".")
If dotPoint > 0 Then
  ExtractFileExt = Right$(strUp, Len(strUp) - dotPoint)
Else
  ExtractFileExt = ""
End If
  
End Function

Private Sub ProcessFiles(ByVal path As String)
'
' Обходим подкаталоги в поисках LNK файлов

Dim objName As String
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

Dim FileName As String
Dim FileExt As String
Dim LnkPath As String

Dim DoEventsCicle As Long

Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
  Do While Cont
    If StopSearch Then Exit Sub
    objName = Left$(WFD.cFileName, InStr(WFD.cFileName, Chr$(0)) - 1)
    If objName <> "." And objName <> ".." Then
      If GetFileAttributes(path & objName) = FILE_ATTRIBUTE_DIRECTORY Then    'Ведь без этого она не сможет отличить файл от папки?
        ProcessFiles path & objName & "\"
      Else
        ' Обработка файлов
        FilesCount = FilesCount + 1
        FileName = objName
        FileExt = ExtractFileExt(objName)
        If FileExt = "LNK" Then
          LnkCount = LnkCount + 1
          LnkPath = GetLinkPath(path & FileName)
          If Not FileExists(LnkPath) Then
            BadLnk = BadLnk + 1
            List1.AddItem (path & FileName)
          End If
        End If
      End If
    End If
    Cont = FindNextFile(hSearch, WFD)
    Label2.Caption = "Обработано " & FilesCount & " файлов, ярлыков " & LnkCount & ", из них битых " & BadLnk
    ' Обновление интерфейса раз в 10 файлов
    DoEventsCicle = DoEventsCicle + 1
    If DoEventsCicle >= 10 Then
      DoEventsCicle = 0
      DoEvents
    End If
  Loop
  Cont = FindClose(hSearch)
End If

End Sub

Private Sub cmdStartStop_Click()
'
' Обработка

If Right$(Text1.Text, 1) <> "\" Then Text1.Text = Text1.Text & "\"

If SearchState Then
  StopSearch = True
  cmdStartStop.Caption = "Старт"
Else
  StopSearch = False
  FilesCount = 0
  LnkCount = 0
  BadLnk = 0
  List1.Clear
  cmdStartStop.Caption = "Стоп"
  SearchState = True
  ProcessFiles Text1.Text
  SearchState = False
  cmdStartStop.Caption = "Старт"
  Label2.Caption = Label2.Caption & ". Завершено."
End If

End Sub

Private Sub Command2_Click()
'
' Открыть каталог с ярлыком и выделить его

Dim LnkFile As String

LnkFile = List1.List(List1.ListIndex)
If FileExists(LnkFile) Then ShellExecute frmMain.hwnd, "OPEN", "EXPLORER", "/select, " & LnkFile, 0, SW_SHOWNORMAL
  
End Sub

Private Sub Command3_Click()
'
' Вносим изменения в ярлык

Dim LnkFile As String

LnkFile = List1.List(List1.ListIndex)
If FileExists(LnkFile) Then
  On Error Resume Next
    With CreateObject("Wscript.Shell").CreateShortcut(LnkFile)
      .TargetPath = Text2.Text
      .Save
      .Close
    End With
End If

End Sub

Private Sub Command4_Click()
'
' Перебираем последовательно вверх каталоги, куда ссылается ярлык и открываем ближайший существующий

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

Private Sub List1_Click()
'
' Клик по пункту списка - по битому ярлыку

Dim LnkFile As String
Dim LnkPath As String

LnkFile = List1.List(List1.ListIndex)
If FileExists(LnkFile) Then
  LnkPath = GetLinkPath(LnkFile)
  Text2.Text = LnkPath
End If

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
'
' Нажатие Del на списке, удаление файлов ярлыков

Dim OneLnk As String
Dim LnkCount As Long
Dim MustDel As Long
Dim i As Long

If KeyCode = 46 Then
  ' Считаем количество выделенных файлов
  If List1.SelCount = 1 Then
    OneLnk = List1.List(List1.ListIndex)
    MustDel = MsgBox("Удалить файл ярлыка:" & vbCrLf & OneLnk, vbYesNo, "Удалить ярлык")
  Else
    LnkCount = List1.SelCount
    MustDel = MsgBox("Выделено " & LnkCount & " файлов, удалить?", vbYesNo, "Удалить ярлык")
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

Sub Form_Resize()
'
' Масштабирование окна

Text1.Width = frmMain.Width - 2415
List1.Width = frmMain.Width - 495
Command2.Left = frmMain.Width - 5910
Command4.Left = frmMain.Width - 3750
Text2.Width = frmMain.Width - 1695
Command3.Left = frmMain.Width - 1470
List1.Height = frmMain.Height - 2760
Text2.Top = frmMain.Height - 1140
Command3.Top = frmMain.Height - 1140

End Sub

