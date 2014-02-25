VERSION 5.00
Begin VB.Form frmQuery 
   Caption         =   "������������� ���������� �������"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   450
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
   ScaleHeight     =   4605
   ScaleWidth      =   9765
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox txtQueryMessage 
      Appearance      =   0  '������
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '���
      Height          =   3735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   9495
   End
   Begin VB.CommandButton cmdMode2 
      Caption         =   "�������� ������"
      Height          =   360
      Left            =   3975
      TabIndex        =   2
      Top             =   4080
      Width           =   2430
   End
   Begin VB.CommandButton cmdMode1 
      Caption         =   "�������� ����� �����"
      Height          =   360
      Left            =   1455
      TabIndex        =   1
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "������"
      Default         =   -1  'True
      Height          =   360
      Left            =   6495
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
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
' ��������� �����, � ������ 0- Cancel, 1- �������� ����� �����, 2- ��������
' ������ ���������

ExitCode = 0
txtQueryMessage.Text = Prompt
'frmQuery.ModeCancel.SetFocus ' �������� VB, ��� �������� ���������???
frmQuery.Show 1, frmMain
QueryMode = ExitCode

End Function

Private Sub cmdMode1_Click()
'
' ������ ����� 1 - ���������� ����� ����� ������ � �������

ExitCode = 1
frmQuery.Hide
 
End Sub

Private Sub cmdMode2_Click()
'
' ������ ����� 2 - ����� ������ ������ � ������� �� �����

ExitCode = 2
frmQuery.Hide
 
End Sub

Private Sub cmdCancel_Click()
'
' ������ ���������

ExitCode = 0
frmQuery.Hide
 
End Sub

Sub Form_Resize()
'
' ��������������� ����

If Me.WindowState <> 1 Then
  txtQueryMessage.Width = frmQuery.Width - 490
  txtQueryMessage.Height = frmQuery.Height - 1320
  cmdMode1.Top = frmQuery.Height - 1070
  cmdMode2.Top = frmQuery.Height - 1070
  cmdCancel.Top = frmQuery.Height - 1070
  cmdMode1.Left = (frmQuery.Width / 2) - (cmdMode2.Width / 2) - cmdMode1.Width - 100
  cmdMode2.Left = (frmQuery.Width / 2) - (cmdMode2.Width / 2)
  cmdCancel.Left = (frmQuery.Width / 2) + (cmdMode2.Width / 2) + 100
  
End If

End Sub
