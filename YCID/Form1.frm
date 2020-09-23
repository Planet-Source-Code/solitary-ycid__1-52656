VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Password"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "User ID"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#############################################
'## This Code Made By Mehran Talaie in 2004 ##
'#############################################


Public Function WindowsDirectory() As String
Dim WinPath As String
Dim temp
WinPath = String(145, Chr(0))
temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function
Private Sub Command1_Click()
MsgBox "This Program made by mehran talaie in 2004", vbOKOnly, "About Show Pass 2.5"
End Sub

Private Sub Command2_Click()
Dim gpwrd As New YCrypto
Set gpwrd = New YCrypto

Dim username As String
gpwrd2 = ReadKey("HKEY_CURRENT_USER\Software\Yahoo\Pager\EOptions String")
username = ReadKey("HKEY_CURRENT_USER\Software\Yahoo\Pager\Yahoo! User ID")
gpwrd.Init 1, 0, username
pass = gpwrd.Decrypt(gpwrd2)
Text2.Text = pass
Text1.Text = ReadKey("HKEY_CURRENT_USER\Software\Yahoo\Pager\Yahoo! User ID")

End Sub

Private Sub Command3_Click()
Dim gpwrd As New YCrypto
Set gpwrd = New YCrypto

Dim username As String
gpwrd2 = ReadKey("HKEY_CURRENT_USER\Software\Yahoo\Pager\EOptions String")
username = ReadKey("HKEY_CURRENT_USER\Software\Yahoo\Pager\Yahoo! User ID")
gpwrd.Init 1, 0, username

pass = gpwrd.Encrypt(gpwrd2)
Text2.Text = pass
End Sub

Private Sub Command4_Click()
End
End Sub



