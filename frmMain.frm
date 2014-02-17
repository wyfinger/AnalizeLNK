VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Проверка ярлыков"
   ClientHeight    =   10170
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   16710
   StartUpPosition =   3  'Windows Default
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
      Width           =   7095
   End
   Begin VB.Menu pm_List 
      Caption         =   "й"
      Visible         =   0   'False
      Begin VB.Menu mi_OpenLnkFile 
         Caption         =   "Открыть каталог с файлом ярлыка"
      End
      Begin VB.Menu mi_OpenFarFolder 
         Caption         =   "Отклыть ближайший существующий каталог"
      End
      Begin VB.Menu mi_SaveList 
         Caption         =   "Сохранить список ярлыков с файл"
         Begin VB.Menu mi_SaveLnkFilesList 
            Caption         =   "Список Lnk файлв"
         End
         Begin VB.Menu mi_SaveDestFilesList 
            Caption         =   "Список целей ярлыков"
         End
         Begin VB.Menu mi_SaveBothList 
            Caption         =   "Lnk фал <-> цель"
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal path As String) As Boolean
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Const SW_SHOWNORMAL = 1

Dim StopSearch As Boolean ' остановить поиск
Dim SearchState As Boolean ' True - работаем, False - не работаем
Dim FilesCount As Long
Dim LnkCount As Long
Dim BadLnk As Long
Dim BothPart As String
Dim InitialDir As String

Private Function GetLinkPath(ByVal lnk As String) As String
'
' Получение пути, на который ссылается ярлык

GetLinkPath = ""
On Error Resume Next
  With CreateObject("Wscript.Shell").CreateShortcut(lnk)
    GetLinkPath = .TargetPath
    .Close
  End With
  
End Function

Private Sub SetLinkPath(ByVal lnkFile As String, ByVal lnk As String)
'
' Изменение пути, на который ссылается ярлык

On Error Resume Next
  With CreateObject("Wscript.Shell").CreateShortcut(lnkFile)
    .TargetPath = Text2.Text
    .Save
    .Close
  End With
  
End Sub

Private Function IsRealyLnk(ByVal lnkFile As String) As Boolean
'
' Проверка действительно ли этот файл является Shell Link (.LNK) Binary File
' См. http://msdn.microsoft.com/en-us/library/dd871305.aspx

Dim SFile As Integer
Dim Readed4 As Long

SFile = FreeFile
Open lnkFile For Binary Access Read As SFile
Get SFile, , Readed4
' Первые 4 байта должны быть 0x0000004C
If Readed4 <> 76 Then Exit Function
' Следующие 16 байт должны быть 00021401-0000-0000-C000-000000000046
Get SFile, , Readed4: If Readed4 <> 136193 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 0 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 192 Then Exit Function
Get SFile, , Readed4: If Readed4 <> 1174405120 Then Exit Function
Close SFile

IsRealyLnk = True

Exit Function
err:
  ' Нет доступа или еще чего-нибудь, не важно, для нас это не ярлык
  IsRealyLnk = False
End Function

Private Sub Command1_Click()
 MsgBox IsRealyLnk("\\Prim-fs-serv\rdu\СРЗА\ТКЗ\Линии\Архив\ЛЭП 110 кВ ВТЭЦ-1 - Орлиная (перспектива)\Программма Параметр\ВЛ 110 кВ А-Голубинка-Орлиная-ВТЭЦ-1 исх.lnk")
End Sub

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

Private Function ExtractFilePath(ByVal strPath As String) As String
'
' Получение каталога файла

Dim slash_pos As Long

slash_pos = InStrRev(Replace$(strPath, "/", "\"), "\")
If slash_pos > 0 Then
  ExtractFilePath = Left$(strPath, Len(strPath) - slash_pos)
Else
  ExtractFilePath = ""
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
        ' Обработка файлов
        FilesCount = FilesCount + 1
        FileName = objName
        FileExt = ExtractFileExt(objName)
        If (FileExt = "LNK") Then
          If IsRealyLnk(path & FileName) Then  ' В VB6 нет частичного вычисления!
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
    Label2.Caption = "Обработано " & FilesCount & " файлов, ярлыков " & LnkCount & ", из них битых " & BadLnk
    ' Обновление интерфейса раз в 10 файлов
    DoEventsCycle = DoEventsCycle + 1
    If DoEventsCycle >= 10 Then
      DoEventsCycle = 0
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

Private Sub mi_OpenLnkFile_Click()
'
' Открыть каталог с ярлыком и выделить его

Dim lnkFile As String

lnkFile = List1.List(List1.ListIndex)
If FileExists(lnkFile) Then ShellExecute frmMain.hwnd, "OPEN", "EXPLORER", "/select, " & lnkFile, 0, SW_SHOWNORMAL
  
End Sub

Private Sub Command3_Click()
'
' Вносим изменения в ярлык

Dim lnkFile As String
Dim lnk As String
Dim i As Long

' Если был выделен один файл - просто изменим его ссылку,
' если выделено несколько файлов - спросим у пользователя как поступить:
' 1. изменить у всех выделенных ярлыков общую часть в пути
' 2. изменить у всех выделенных ярлыков всю ссылку нановый (когда несколько
' ярлыков ссылаются на один файл

If List1.SelCount = 1 Then
  lnkFile = List1.List(List1.ListIndex)
  If FileExists(lnkFile) Then
    SetLinkPath lnkFile, Text2.Text
  End If
Else
  ' Подготовим информативное сообщение для пользователя, чтобы не запутался
  Select Case frmQuery.QueryMode("У объектов, на которые ссылаются выделенные Вами ярлыки имеется " & _
  "следующая общая часть:" & vbCrLf & BothPart & vbCrLf & vbCrLf & _
  "Заменить в выделеных ярлыках общую часть пути на новую, или просто обновить ссылку?")
    Case 0
      ' Ничего не делаем
    Case 1
      ' Замена общей части в путях
      For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          lnkFile = List1.List(i)
          lnk = GetLinkPath(lnkFile)
          lnk = Text2.Text & Mid$(lnk, Len(BothPart) + 1, Len(lnk) - Len(BothPart) - 1)
          SetLinkPath lnkFile, lnk
        End If
      Next
    Case 2
      ' Просто замена путей на новые
      For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
          lnkFile = List1.List(i)
          SetLinkPath lnkFile, Text2.Text
        End If
      Next
  End Select
End If

End Sub

Private Sub mi_OpenFarFolder_Click()
'
' Перебираем последовательно вверх каталоги, куда ссылается ярлык и открываем ближайший существующий

Dim lnkFile As String
Dim LnkPath As String
Dim SlashPos As Long

lnkFile = List1.List(List1.ListIndex)
If FileExists(lnkFile) Then
  LnkPath = GetLinkPath(lnkFile)
    
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

Dim SelectedLinks As New Collection
Dim lnkFile As String
Dim LnkPath As String
Dim i As Long
Dim j As Long
Dim p As Long
Dim CharA As String
Dim CharB As String
Dim FindDiff As Boolean

' Если выделен один ярлык - просто отобразим ссылку из ярлыка
' в противном случае выделим общую часть BothPart
If List1.SelCount = 1 Then
  BothPart = ""
  lnkFile = List1.List(List1.ListIndex)
  If FileExists(lnkFile) Then
    LnkPath = GetLinkPath(lnkFile)
    Text2.Text = LnkPath
  End If
Else
  For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then
      SelectedLinks.Add GetLinkPath(List1.List(i))
    End If
  Next
  ' Ищем общую часть
  p = 1
  FindDiff = False
  Do While Not FindDiff
    ' Берем очередной символ из первой строки, и сверяем с остальными
    CharA = Mid$(SelectedLinks(1), p, 1)
    For j = 2 To List1.SelCount
      CharB = Mid$(SelectedLinks(j), p, 1)
      If CharA <> CharB Then
        FindDiff = True
        Exit For
      End If
    Next
    p = p + 1
  Loop
  BothPart = Mid$(SelectedLinks(1), 1, p - 2)
  ' Отсечем символы до первой "\" в конце строки на случай, если начала имен
  ' файлов или каталогов совпадают
  BothPart = Left$(BothPart, InStrRev(BothPart, "\"))
  Text2.Text = BothPart
End If

Set SelectedLinks = Nothing

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
' Контекстое меню на списке файлов


If Button = 2 Then
  ' Выделим элемент при клике правой кновкой мыши
  ' Сделаю это потом
  ' List1.TopIndex
 
  If (List1.SelCount = 1) And (FileExists(List1.List(List1.TabIndex))) Then
    mi_OpenLnkFile.Enabled = True
    mi_OpenFarFolder.Enabled = True
  Else
    mi_OpenLnkFile.Enabled = False
    mi_OpenFarFolder.Enabled = False
  End If
  If List1.ListCount > 0 Then
    mi_SaveList.Enabled = True
  Else
    mi_SaveList.Enabled = False
  End If
  PopupMenu pm_List
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
Text2.Width = frmMain.Width - 1695
Command3.Left = frmMain.Width - 1470
List1.Height = frmMain.Height - 2760
Text2.Top = frmMain.Height - 1140
Command3.Top = frmMain.Height - 1140
Line1.X2 = frmMain.Width + 5000

End Sub

Private Function SaveFileDialog(InitialDir As String) As String
'
' Вызываем диалог выбора файла для сохраненния

Dim OFName As OPENFILENAME
If InitialDir = "" Then InitialDir = App.path

OFName.lStructSize = Len(OFName)
OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Все файлв (*.*)" + Chr$(0) + "*.*" + Chr$(0)
OFName.lpstrFile = Space$(254)
OFName.nMaxFile = 255
OFName.lpstrFileTitle = Space$(254)
OFName.nMaxFileTitle = 255
OFName.lpstrInitialDir = InitialDir
OFName.lpstrTitle = "Сохранение файла"
OFName.flags = 0
 
If GetSaveFileName(OFName) Then
  SaveFileDialog = Trim$(OFName.lpstrFile)
Else
  SaveFileDialog = ""
End If

End Function

Sub mi_SaveLnkFilesList_Click()
'
' Сохранение списка LNK файлов

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

Sub mi_SaveDestFilesList_Click()
'
' Сохранение списка целей ярлыков

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

Sub mi_SaveBothList_Click()
'
' Сохранение списка LNK файлов и их целей

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
