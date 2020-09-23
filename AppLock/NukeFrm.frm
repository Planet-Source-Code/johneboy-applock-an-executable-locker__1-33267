VERSION 5.00
Begin VB.Form NukeFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NukeAppLock"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "NukeFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Text            =   """%1"" %*"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "NukeFrm.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to remove AppLock and all the components?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "NukeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\DefaultIcon"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\InProcServer32"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\Shell\Open\Command"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\ShellEx\PropertySheetHandlers\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
DeleteKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\ShellFolder"
DeleteKey "HKEY_CURRENT_USER\Software\AppLock"
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
DeleteKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\open\command", "", "" + Text1.Text + ""
SetStringValue "HKEY_CLASSES_ROOT\exefile\shell\runas\command", "", "" + Text1.Text + ""
MsgBox "AppLock has been nuked!", vbOKOnly + vbInformation, "AppLock Nuked"
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
