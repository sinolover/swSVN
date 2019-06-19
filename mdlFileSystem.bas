Attribute VB_Name = "mdlFileSystem"
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Const MAX_LEN = 200 '字符串最大长度
Const DESKTOP = &H0& '桌面
Const PROGRAMS = &H2& '程序集
Const MYDOCUMENTS = &H5& '我的文档
Const MYFAVORITES = &H6& '收藏夹
Const STARTUP = &H7& '启动
Const RECENT = &H8& '最近打开的文件
Const SENDTO = &H9& '发送
Const STARTMENU = &HB& '开始菜单
Const NETHOOD = &H13& '网上邻居
Const FONTS = &H14& '字体
Const SHELLNEW = &H15& 'ShellNew
Const APPDATA = &H1A& 'Application Data
Const PRINTHOOD = &H1B& 'PrintHood
Const PAGETMP = &H20& '网页临时文件
Const COOKIES = &H21& 'Cookies目录
Const HISTORY = &H22& '历史
Public Function getSpecialFolder(va As String) As String
Dim sTmp As String * MAX_LEN   '存放结果的固定长度的字符串
Dim nLength As Long   '字符串的实际长度
Dim pidl As Long   '某特殊目录在特殊目录列表中的位置
va = LCase(Trim(va))
If Len(va) < 2 Then va = "windows"
Select Case va
Case "windows"
'*************************获得Windows目录**********************************
nLength = GetWindowsDirectory(sTmp, MAX_LEN)
getSpecialFolder = Left(sTmp, nLength)
Case "system"
'*************************获得System目录***********************************
nLength = GetSystemDirectory(sTmp, MAX_LEN)
getSpecialFolder = Left(sTmp, nLength)
'*************************获得Temp目录***********************************
Case "temp"
nLength = GetTempPath(MAX_LEN, sTmp)
getSpecialFolder = Left(sTmp, nLength)
'*************************获得DeskTop目录**********************************
Case "desktop"
SHGetSpecialFolderLocation 0, DESKTOP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得发送到目录**********************************
Case "sendto"
SHGetSpecialFolderLocation 0, SENDTO, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得我的文档目录*********************************
Case "mydocuments", "documents", "document"
SHGetSpecialFolderLocation 0, MYDOCUMENTS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得程序集目录***********************************
Case "programs", "program", "program files"
SHGetSpecialFolderLocation 0, PROGRAMS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得启动目录*************************************
Case "startup"
SHGetSpecialFolderLocation 0, STARTUP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得开始菜单目录*********************************
Case "startmenu"
SHGetSpecialFolderLocation 0, STARTMENU, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)

'*************************获得收藏夹目录***********************************
Case "myfavorites", "favorites"
SHGetSpecialFolderLocation 0, MYFAVORITES, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'**********************获得最后打开的文件目录*******************************
Case "recent"
SHGetSpecialFolderLocation 0, RECENT, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得网上邻居目录*********************************
Case "nethood"
SHGetSpecialFolderLocation 0, NETHOOD, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得字体目录**********************************
Case "fonts"
SHGetSpecialFolderLocation 0, FONTS, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得Cookies目录**********************************
Case "cookies"
SHGetSpecialFolderLocation 0, COOKIES, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得历史目录**********************************
Case "history", "histories"
SHGetSpecialFolderLocation 0, HISTORY, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'***********************获得网页临时文件目录*******************************
Case "pagetmp"
SHGetSpecialFolderLocation 0, PAGETMP, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得ShellNew目录*********************************
Case "shellnew"
SHGetSpecialFolderLocation 0, SHELLNEW, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'***********************获得Application Data目录*****************************
Case "appdata"
SHGetSpecialFolderLocation 0, APPDATA, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
'*************************获得PrintHood目录*********************************
SHGetSpecialFolderLocation 0, PRINTHOOD, pidl
SHGetPathFromIDList pidl, sTmp
getSpecialFolder = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
End Select
End Function

