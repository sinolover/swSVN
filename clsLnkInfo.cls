VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLnkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'注意要引用：Microsoft Shell Controls And Automation
Private Sub Command1_Click()
    Dim FolderPath As String
    Dim ShortcutName As String
    Dim WorkDir As String
    Dim Arguments As String, Description As String
    Dim IconIdx As Long, ShowCommand As Long
    FolderPath = "e:\mc\"
    ShortcutName = "大补贴.lnk"
    Dim IconFile As String
    Call GetShellLinkInfo(FolderPath, ShortcutName, WorkDir, Arguments, Description, IconFile, IconIdx, ShowCommand)
   
End Sub


Private Sub GetShellLinkInfo(ByVal FolderPath As String, ByVal ShortcutName As String, WorkDir As String, _
                                  Arguments As String, Description As String, IconFile As String, IconIdx As Long, _
                                 ShowCommand As Long)
   Dim mShell As Shell, mFile As FolderItem, mFolder As Folder
   Dim lnk As ShellLinkObject, i As Long
    
   Set mShell = New Shell
   Set mFolder = mShell.NameSpace(FolderPath)
   On Error Resume Next
   Set mFile = mFolder.Items.Item(ShortcutName)
   If Err Then
      MsgBox ShortcutName & " is inaccessable!"
      Err.Clear
      GoTo exit_sub
   Else
      If mFile.IsLink Then
          Set lnk = mFile.GetLink
          WorkDir = lnk.WorkingDirectory
          Arguments = lnk.Arguments
          Description = lnk.Description
          IconIdx = lnk.GetIconLocation(IconFile)
          ShowCommand = lnk.ShowCommand
          MsgBox "Name: " & mFile.Name & vbCrLf & _
          "Description: " & lnk.Description & vbCrLf & _
          "Path: " & lnk.Path & vbCrLf & _
          "WorkingDirectory: " & lnk.WorkingDirectory & vbCrLf, vbInformation
      Else
         MsgBox ShortcutName & " is not a shortcut!", vbInformation
      End If
   End If
exit_sub:
   Set lnk = Nothing
   Set mFile = Nothing
   Set mFolder = Nothing
   Set mShell = Nothing
End Sub

