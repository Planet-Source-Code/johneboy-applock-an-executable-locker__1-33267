VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AppLock"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "Form2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2865
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2865
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   3975
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
         Top             =   1320
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password to lock and unlock executables.  Please remember this password.  Once it is set, there is no way to recover it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form2.frx":0CCA
      Top             =   2760
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If Text1 = Text2 Then
    SetStringValue "HKEY_CURRENT_USER\Software\AppLock", "FirstRun", "No"
    SetStringValue "HKEY_CURRENT_USER\Software\AppLock", "Password", "" + Text3.Text + ""
    Dim LockNow As Variant
    LockNow = MsgBox("Do you wish to lock all executables now?", vbYesNo + vbQuestion, "Lock Now?")
    If LockNow = vbYes Then
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\open\command", "", "" & GetAppPath
    SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\runas\command", "", "" & GetAppPath
    MsgBox "All executables have been locked!", vbOKOnly + vbInformation, "EXE's Locked!"
    Unload Me
    Else
    Unload Me
    End If
Else
    MsgBox "Confirmed password does not match. Please retype.", vbOKOnly + vbCritical, "Password Error"
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
SetStringValue "HKEY_CURRENT_USER\Software\AppLock", "FirstRun", "Yes"
Unload Me
End Sub
Private Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function
Private Function TEncrypt(iString)
Dim q As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As String
Dim x As Variant
Dim f As Variant
On Error GoTo uhoh
    q = ""
    a = randomnumber(9) + 32
    b = randomnumber(9) + 32
    c = randomnumber(9) + 32
    d = randomnumber(9) + 32
    q = Chr(a) & Chr(c) & Chr(b)
    e = 1
    For x = 1 To Len(iString)
        f = Mid(iString, x, 1)
        If e = 1 Then q = q & Chr(Asc(f) + a)
        If e = 2 Then q = q & Chr(Asc(f) + c)
        If e = 3 Then q = q & Chr(Asc(f) + b)
        If e = 4 Then q = q & Chr(Asc(f) + d)
        e = e + 1
        If e > 4 Then e = 1
    Next x
    q = q & Chr(d)
    TEncrypt = q
    Exit Function
uhoh:
    TEncrypt = "Error: Invalid text To Encrypt"
    Exit Function
End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Image1_Click()
Form4.Show
End Sub

Private Sub Text1_Change()
Text1.Text = Replace(Text1.Text, " ", "")
Text3 = TEncrypt(Text1)
If Text1 = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Call Command1_Click
End If

End Sub
