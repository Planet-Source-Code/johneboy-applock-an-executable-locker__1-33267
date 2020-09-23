Attribute VB_Name = "SubMain"
Option Explicit
Public Sub Main()
Dim EXEpath As String
Dim Path As String
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"


If FirstRun = "No" Then
SetStringValue "HKEY_CURRENT_USER\Software\AppLock", "AppPath", "" & EXEpath
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe,0"
SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\Shell\Open\Command", "", "" & EXEpath
    Path = Command
        If Path = "" Then
        Form3.Show
        Else
        Form1.Show
        End If

Else
CreateKey "HKEY_CURRENT_USER\Software\AppLock"
SetStringValue "HKEY_CURRENT_USER\Software\AppLock", "AppPath", "" & EXEpath

'{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}
' the GUID i will use for the app

 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\DefaultIcon"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\InProcServer32"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\Shell\Open\Command"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\ShellEx\PropertySheetHandlers\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
 CreateKey "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\ShellFolder"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}", "", "AppLock"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}", "InfoTip", "AppLock"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\DefaultIcon", "", "" + App.Path + "\" + App.EXEName + ".exe,0"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\InProcServer32", "", "Shell32.dll"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\InProcServer32", "ThreadingModel", "Apartment"
 SetStringValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\Shell\Open\Command", "", "" & EXEpath
 CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
 CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}"
 SetBinaryValue "HKEY_CLASSES_ROOT\CLSID\{8C6D8BD6-116B-4D4E-B1C2-87098DB509BB}\ShellFolder", "Attributes", Chr$(&H0) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)

Form2.Show
End If
End Sub
