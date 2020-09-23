VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AppLock"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "Form3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3600
      TabIndex        =   9
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Text            =   """%1"" %*"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Currently Unlocked"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password to Unlock/Lock all Executables."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1905
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1905
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form3.frx":0CCA
      Top             =   1800
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
If Text1 = Text2 Then
    If Label2.Caption = "Currently Locked" Then
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\open\command", "", "" + Text3.Text + ""
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\runas\command", "", "" + Text3.Text + ""
    MsgBox "Executables have been unlocked", vbOKOnly + vbInformation, "Executables Unlocked"
    Unload Me
    Else
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\open\command", "", "" & GetAppPath
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\runas\command", "", "" & GetAppPath
    MsgBox "Executables have been Locked", vbOKOnly + vbInformation, "Executables Locked"
    Unload Me
    End If
Else
MsgBox "Incorrect Password! Executables will not be unlocked.", vbOKOnly + vbCritical, "Incorrect Password!"
Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Function TDecrypt(iString)
Dim q As String
Dim zz As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim x As Variant
Dim f As Variant
Dim txt As String
Dim txt2 As String
    On Error GoTo uhohs
    q = ""
    zz = Left(iString, 3)
    a = Left(zz, 1)
    b = Mid(zz, 2, 1)
    c = Mid(zz, 3, 1)
    d = Right(iString, 1)
    a = Int(Asc(a))
    b = Int(Asc(b))
    c = Int(Asc(c))
    d = Int(Asc(d))
    txt = Left(iString, Len(iString) - 1)
    txt2 = Mid(txt, 4, Len(txt))
    e = 1
    For x = 1 To Len(txt2)
        f = Mid(txt2, x, 1)
        If e = 1 Then q = q & Chr(Asc(f) - a)
        If e = 2 Then q = q & Chr(Asc(f) - b)
        If e = 3 Then q = q & Chr(Asc(f) - c)
        If e = 4 Then q = q & Chr(Asc(f) - d)
        e = e + 1
        If e > 4 Then e = 1
    Next x
    TDecrypt = q
    Exit Function
uhohs:
    TDecrypt = "Error: Invalid text To Decrypt"
    Exit Function
End Function
Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function
Private Sub Form_Load()
If GetLockStatus = "" + Text3.Text + "" Then
Label2.Caption = "Currently Unlocked"
Else
Label2.Caption = "Currently Locked"
End If
Text4 = GetPassword
Text2 = TDecrypt(Text4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Form4.Show
End Sub

Private Sub Text1_Change()
If Text1 = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call Command1_Click
End If

End Sub
