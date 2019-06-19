VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "移动办公地址切换"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4650
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "置VPN"
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "启用位置"
      Height          =   1815
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   3015
      Begin VB.OptionButton Op0 
         Caption         =   "未设办公地址网络"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Op2 
         Caption         =   "外出移动办公网络"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Op1 
         Caption         =   "公司内部办公网络"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   435
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设定"
      Default         =   -1  'True
      Height          =   435
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cstFileName As String


Private Sub Command2_Click() '404990
Unload Me
End Sub

Private Sub Form_Load()
Dim strPath

On Error Resume Next
''注意要引用：Microsoft Shell Controls And Automation
'Dim FolderPath As String
'Dim ShortcutName As String
'Dim WorkDir As String
'Dim Arguments As String, Description As String
'
'Dim IconFile As String
'
'Dim IconIdx As Long, ShowCommand As Long
'Dim Link As String



Dim objLink As Object
Dim objScript As Object


strPath = getSpecialFolder("windows")
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
cstFileName = strPath & "system32\drivers\etc\hosts"
If Len(cstFileName) < 28 Then cstFileName = "%windir%\system32\drivers\etc\hosts"
 
 

strPath = getSpecialFolder("desktop")
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
strPath = strPath & "OpenVPN GUI.lnk"

If Dir(strPath) = "" Then strPath = "%homepath%\desktop\OpenVPN GUI.lnk"


If Dir(strPath) <> "" Then
  'If GetShellLinkInfo(FolderPath, ShortcutName, WorkDir, Arguments, Description, IconFile, IconIdx, ShowCommand, Link) Then
    'If InStr(Arguments, "--connect") < 1 Then
    
  Set objScript = CreateObject("WScript.Shell")
      
  Set objLink = objScript.CreateShortcut(strPath)
      
  With objLink
      '.TargetPath = .TargetPath & " --connect client.ovpn --silent_connection 1"          '目标
      .Arguments = " --connect client.ovpn --silent_connection 1"
      .Save                       '保存
  
  '.Hotkey = "Ctrl + Alt + E"  '快捷键
  '.WindowStyle = 1            '
  '.IconLocation = ""          '图标
  '.Description = "描述"       '描述
  '.WorkingDirectory = strPath '起始位置
  
  End With
    'End If
  'End If
End If



strPath = getSpecialFolder("allusersdesktop")
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
strPath = strPath & "OpenVPN GUI.lnk"

If Dir(strPath) = "" Then strPath = "%public%\desktop\OpenVPN GUI.lnk"


If Dir(strPath) <> "" Then
  'If GetShellLinkInfo(FolderPath, ShortcutName, WorkDir, Arguments, Description, IconFile, IconIdx, ShowCommand, Link) Then
    'If InStr(Arguments, "--connect") < 1 Then
    
  Set objScript = CreateObject("WScript.Shell")
      
  Set objLink = objScript.CreateShortcut(strPath)
      
  With objLink
      '.TargetPath = .TargetPath & " --connect client.ovpn --silent_connection 1"          '目标
      .Arguments = " --connect client.ovpn --silent_connection 1"
      .Save                       '保存
  
  '.Hotkey = "Ctrl + Alt + E"  '快捷键
  '.WindowStyle = 1            '
  '.IconLocation = ""          '图标
  '.Description = "描述"       '描述
  '.WorkingDirectory = strPath '起始位置
  
  End With
    'End If
  'End If
End If




strPath = getSpecialFolder("startmenu")
If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
strPath = strPath & "Programs\OpenVPN\OpenVPN GUI.lnk"
If Dir(strPath) = "" Then
strPath = "%systemdrive%\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenVPN\OpenVPN GUI.lnk"
End If
If Dir(strPath) Then
  Set objScript = CreateObject("WScript.Shell")
      
  Set objLink = objScript.CreateShortcut(strPath)
      
  With objLink
      .Arguments = " --connect client.ovpn --silent_connection 1"
      .Save                       '保存
  
  
  End With
End If





Select Case getHosts()
Case 1
  Op1.Value = True
Case 2
  Op2.Value = True
  
Case Else
  Op0.Value = True
End Select
End Sub

Private Function GetShellLinkInfo(ByVal FolderPath As String, ByVal ShortcutName As String, WorkDir As String, _
                                 Arguments As String, Description As String, IconFile As String, IconIdx As Long, _
                                ShowCommand As Long, Link As String)
  Dim mShell As Shell, mFile As FolderItem, mFolder As Folder
  Dim lnk As ShellLinkObject, i As Long
  
  GetShellLinkInfo = False
  Set mShell = New Shell
  Set mFolder = mShell.NameSpace(FolderPath)
  On Error Resume Next
  Set mFile = mFolder.Items.Item(ShortcutName)
  If Not Err Then
    If mFile.IsLink Then
      GetShellLinkInfo = True
      Set lnk = mFile.GetLink
      WorkDir = lnk.WorkingDirectory
      Arguments = lnk.Arguments
      Description = lnk.Description
      IconIdx = lnk.GetIconLocation(IconFile)
      ShowCommand = lnk.ShowCommand
      Link = lnk.Path
'      MsgBox "Name: " & mFile.Name & vbCrLf & _
      "Description: " & lnk.Description & vbCrLf & _
      "Path: " & lnk.Path & vbCrLf & _
      "WorkingDirectory: " & lnk.WorkingDirectory & vbCrLf, vbInformation
    End If
  End If
exit_sub:
  Set lnk = Nothing
  Set mFile = Nothing
  Set mFolder = Nothing
  Set mShell = Nothing
End Function
Private Sub Command1_Click()
Dim rv As Boolean
On Error Resume Next
Err.Clear
rv = False
Call setHosts
Select Case getHosts()
Case 1
 If Op1.Value = True Then MsgBox "修改完成，当前工作位置：" & Op1.Caption, vbInformation, "【办公位置】设置成功"
 rv = True
Case 2
 If Op2.Value = True Then MsgBox "修改完成，当前工作位置：" & Op2.Caption, vbInformation, "【办公位置】设置成功"
 rv = True
Case 0
 If Op0.Value = True Then MsgBox "修改完成，当前工作位置：" & Op0.Caption, vbInformation, "取消【办公位置】成功"
 rv = True
End Select
If rv = False Or Err.Number > 0 Then MsgBox "请右击本工具软件，并选择【以管理员身份运行】，若杀毒软件出现提示，请选择允许，或者将本工具软件加入白名单", vbCritical, "【办公位置】设置失败"
End Sub

Function getHosts() As Integer
Dim str1 As String
Dim v As Integer
getHosts = 0
List1.Clear
Open cstFileName For Input As #1
  While Not EOF(1)
    Line Input #1, str1
    str1 = LCase(Trim(str1))
    If Len(str1) > 0 Then
    If InStr(str1, "svn.boyuanitsm.com") > 0 Then
      getHosts = 2
      If InStr(str1, "172.16.5.50") > 0 Then getHosts = 1
    Else
      List1.AddItem str1
    End If
    End If
  Wend
Close
End Function

Public Sub setHosts()
On Error Resume Next
Dim i As Integer

If Op1.Value = True Then List1.AddItem "172.16.5.50                 svn.boyuanitsm.com"
If Op2.Value = True Then List1.AddItem "192.168.255.1               svn.boyuanitsm.com"

Open cstFileName For Output As #1
  For i = 0 To List1.ListCount - 1
    Print #1, List1.List(i)
  Next i
Close
End Sub

