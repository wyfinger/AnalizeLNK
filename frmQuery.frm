VERSION 5.00
Begin VB.Form frmQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Корректировка нескольких ярлыков"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9765
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Mode2 
      Caption         =   "Обновить ссылку"
      Height          =   360
      Left            =   3975
      TabIndex        =   3
      Top             =   4080
      Width           =   2430
   End
   Begin VB.CommandButton Mode1 
      Caption         =   "Изменить общую часть"
      Height          =   360
      Left            =   1455
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton ModeCancel 
      Caption         =   "Отмена"
      Default         =   -1  'True
      Height          =   360
      Left            =   6495
      TabIndex        =   1
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label QueryText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Прозрачно
      Caption         =   "Label1"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExitCode As Long

Public Function QueryMode(Prompt As String) As Integer
'
' Отобразим форму, и вернем 0- Cancel, 1- Изменить общую часть, 2- Заменить
' ссылку полностью

ExitCode = 0
frmQuery.QueryText.Caption = Prompt
'frmQuery.ModeCancel.SetFocus ' добанный VB, как выделить компеонент???
frmQuery.Show 1, frmMain
QueryMode = ExitCode

End Function


Private Sub Mode1_Click()
 ExitCode = 1
 frmQuery.Hide
 
End Sub

Private Sub Mode2_Click()
 ExitCode = 2
 frmQuery.Hide
 
End Sub

Private Sub ModeCancel_Click()
 ExitCode = 0
 frmQuery.Hide
 
End Sub

