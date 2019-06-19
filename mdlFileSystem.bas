Attribute VB_Name = "mdlFileSystem"
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Const MAX_LEN = 200 '�ַ�����󳤶�
Const DESKTOP = &H0& '����
Const PROGRAMS = &H2& '����
Const MYDOCUMENTS = &H5& '�ҵ��ĵ�
Const MYFAVORITES = &H6& '�ղؼ�
Const STARTUP = &H7& '����
Const RECENT = &H8& '����򿪵��ļ�
Const SENDTO = &H9& '����
Const STARTMENU = &HB& '��ʼ�˵�
Const NETHOOD = &H13& '�����ھ�
Const FONTS = &H14& '����
Const SHELLNEW = &H15& 'ShellNew
Const APPDATA = &H1A& 'Application Data
Const PRINTHOOD = &H1B& 'PrintHood
Const PAGETMP = &H20& '��ҳ��ʱ�ļ�
Const COOKIES = &H21& 'CookiesĿ¼
Const HISTORY = &H22& '��ʷ
Public Function getSpecialFolder(va As String) As String
Dim sTmp As String * MAX_LEN   '��Ž���Ĺ̶����ȵ��ַ���
Dim nLength As Long   '�ַ�����ʵ�ʳ���
Dim pidl As Long   'ĳ����Ŀ¼������Ŀ¼�б��е�λ��
va = LCase(Trim(va))
If Len(va) < 2 Then va = "windows"
Select Case va
Case "windows"
'*************************���WindowsĿ¼**********************************
nLength = GetWindowsDirectory(sTmp, MAX_LEN)
getSpecialFolder = Left(sTmp, nLength)
Case "system"
'*************************���SystemĿ¼***********************************
nLength = GetSystemDirectory(sTmp, MAX_LEN)
getSpecialFolder = Left(sTmp, nLength)
'*************************���TempĿ¼***********************************
Case "temp"
nLength = GetTempPath(MAX_LEN, sTmp)
getSpecialFolder = Left(sTmp, nLength)
'*************************���DeskTopĿ¼**********************************
Case "desktop"
SHGetSpecialFolderLocation 0, DESKTOP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************��÷��͵�Ŀ¼**********************************
Case "sendto"
SHGetSpecialFolderLocation 0, SENDTO, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************����ҵ��ĵ�Ŀ¼*********************************
Case "mydocuments", "documents", "document"
SHGetSpecialFolderLocation 0, MYDOCUMENTS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************��ó���Ŀ¼***********************************
Case "programs", "program", "program files"
SHGetSpecialFolderLocation 0, PROGRAMS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************�������Ŀ¼*************************************
Case "startup"
SHGetSpecialFolderLocation 0, STARTUP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************��ÿ�ʼ�˵�Ŀ¼*********************************
Case "startmenu"
SHGetSpecialFolderLocation 0, STARTMENU, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)

'*************************����ղؼ�Ŀ¼***********************************
Case "myfavorites", "favorites"
SHGetSpecialFolderLocation 0, MYFAVORITES, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'**********************������򿪵��ļ�Ŀ¼*******************************
Case "recent"
SHGetSpecialFolderLocation 0, RECENT, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************��������ھ�Ŀ¼*********************************
Case "nethood"
SHGetSpecialFolderLocation 0, NETHOOD, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************�������Ŀ¼**********************************
Case "fonts"
SHGetSpecialFolderLocation 0, FONTS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************���CookiesĿ¼**********************************
Case "cookies"
SHGetSpecialFolderLocation 0, COOKIES, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************�����ʷĿ¼**********************************
Case "history", "histories"
SHGetSpecialFolderLocation 0, HISTORY, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'***********************�����ҳ��ʱ�ļ�Ŀ¼*******************************
Case "pagetmp"
SHGetSpecialFolderLocation 0, PAGETMP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************���ShellNewĿ¼*********************************
Case "shellnew"
SHGetSpecialFolderLocation 0, SHELLNEW, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'***********************���Application DataĿ¼*****************************
Case "appdata"
SHGetSpecialFolderLocation 0, APPDATA, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************���PrintHoodĿ¼*********************************
SHGetSpecialFolderLocation 0, PRINTHOOD, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
End Select
End Function

