VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "U盘防火墙 -- (硬盘区查看版)"
   ClientHeight    =   5370
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7485
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7485
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5400
      Top             =   5040
   End
   Begin VB.FileListBox File1 
      Height          =   4050
      Hidden          =   -1  'True
      Left            =   3360
      MultiSelect     =   2  'Extended
      Pattern         =   "*"
      System          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   3360
      TabIndex        =   5
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   540
   End
   Begin VB.Menu Doing 
      Caption         =   "对文件的操作"
      Visible         =   0   'False
      Begin VB.Menu Run 
         Caption         =   "运行"
      End
      Begin VB.Menu Nothing4 
         Caption         =   "-"
      End
      Begin VB.Menu Delete 
         Caption         =   "删除"
      End
      Begin VB.Menu Copy 
         Caption         =   "复制"
      End
      Begin VB.Menu Move 
         Caption         =   "移动"
      End
      Begin VB.Menu Nothing 
         Caption         =   "-"
      End
      Begin VB.Menu NewFile 
         Caption         =   "刷新"
      End
      Begin VB.Menu Nothing1 
         Caption         =   "-"
      End
      Begin VB.Menu PutOntheProc 
         Caption         =   "用指定的程序打开"
      End
      Begin VB.Menu PutOnNotebook 
         Caption         =   "用记事本打开"
      End
   End
   Begin VB.Menu DoingBack 
      Caption         =   "对文件夹的操作"
      Visible         =   0   'False
      Begin VB.Menu Back 
         Caption         =   "返回"
      End
      Begin VB.Menu OpenBack 
         Caption         =   "打开"
      End
      Begin VB.Menu Create 
         Caption         =   "新建"
      End
      Begin VB.Menu DeleteBack 
         Caption         =   "删除"
      End
      Begin VB.Menu Nothing2 
         Caption         =   "-"
      End
      Begin VB.Menu NewDrive 
         Caption         =   "刷新盘目录"
      End
      Begin VB.Menu NewBack 
         Caption         =   "刷新文件夹"
      End
   End
   Begin VB.Menu File 
      Caption         =   "文件  "
      Begin VB.Menu Scan 
         Caption         =   "扫描所有磁盘"
         Shortcut        =   ^S
      End
      Begin VB.Menu Write 
         Caption         =   "修复漏洞"
         Enabled         =   0   'False
         Shortcut        =   ^M
      End
      Begin VB.Menu Nothing3 
         Caption         =   "-"
      End
      Begin VB.Menu Abuot 
         Caption         =   "关于"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu DriveData 
      Caption         =   "硬盘数据  "
      Begin VB.Menu Look 
         Caption         =   "查看"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetFileType Lib "kernel32" (ByVal hFile As Long) As Long     ''判断该文件属性
Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long      ''打开
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long      ''关闭
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Boolean    '判断文件是否存在于某一个目录
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long     ''判断该磁盘属性
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long     ''删除文件
Private Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long     ''删除文件夹

Private Sub File1_DblClick()       ''''   ( Can doing)   [ But ...it don't start some Proc]
On Error Resume Next
 Dim Name As String
 Name = LCase(Mid(File1.FileName, InStr(File1.FileName, ".") + 1))   ''+1是为了不显示"."号
 If Name = "exe" Or Name = "bat" Then
 Shell File1.Path & "\" & File1.FileName, vbNormalFocus
 ElseIf Name = "txt" Or Name = "dat" Or Name = "ini" Or Name = "inf" Then
 Shell "notepad.exe " + File1.Path & "\" & File1.FileName, vbNormalFocus
 ElseIf Name = "html" Or Name = "htm" Or Name = "xml" Then
 Shell "iexplorer " + File1.Path & "\" & File1.FileName
 Else
 MsgBox "无法运行 " + "." + Name + "文件"
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

'''''               任何还没有经过正确运行的代码放到上面
'''''                    ( Can Doing Datch )

''''  #######################33    菜单代码
''''  #########################
''''  #######################33    菜单 -> 对文件的操作

Private Sub Move_Click()       ''''   ( Can doing)
On Error GoTo a:
If File1.ListIndex = -1 Then MsgBox "还没有在File框中选一个文件移动呢": Exit Sub
Dim Path As String
Path = InputBox("请在下面输入要移动到的路径", "移动文件中", Dir1.Path)
If Path = "" Then MsgBox "请输入一个正确的路径": Exit Sub
If (MsgBox("真的要移动到" + Path + "下吗?", vbQuestion, "移动文件中") = vbOK) Then
Call FileCopy(File1.List(File1.ListIndex), Path & "\" & File1.List(File1.ListIndex))  '''第一个是文件名,最后的是目标文件
Kill (Dir1.Path & "\" & File1.List(File1.ListIndex))    ''利用复制/删除的机理
Label2.Caption = "该文件夹下共有:" & File1.ListCount & "个对象"
File1.Refresh   ''刷新控件   ( 如果不这样弄的话,控件不会自动刷新 )
Exit Sub
End If

a:
MsgBox "路径无效"
End Sub

Private Sub Copy_Click()     ''''   ( Can doing)
On Error GoTo a:
If File1.ListIndex = -1 Then MsgBox "还没有在File框中选一个文件复制呢": Exit Sub
Dim Path As String
Path = InputBox("请在下面输入要复制到的路径", "复制文件中", Dir1.Path)
If Path = "" Then MsgBox "请输入一个正确的路径": Exit Sub
If (MsgBox("真的要复制到" + Path + "  下吗?", vbQuestion, "复制文件中") = vbOK) Then
Call FileCopy(File1.List(File1.ListIndex), Path & "\" & File1.List(File1.ListIndex)) '''第一个是文件名,最后的是目标文件
Exit Sub
End If

a:
MsgBox "路径无效"
End Sub

Private Sub Run_Click()      ''''   ( Can doing)
File1_DblClick
End Sub

Private Sub Delete_Click()      ''''   ( Can doing)
On Error Resume Next
DeleteFile (Dir1.Path & "\" & File1.FileName)
Label2.Caption = "该文件夹下共有:" & File1.ListCount & "个对象"
File1.Refresh    ''刷新控件   ( 如果不这样弄的话,控件不会自动刷新 )
End Sub

Private Sub Create_Click()     ''''   ( Can doing)
Dim Name
Name = InputBox("请输入一个文件名")
If Name = "" Then
MsgBox "无效文件名"
Exit Sub
ElseIf Name = 0 Then
Exit Sub
End If
MkDir (Name)    ''创建一个文件夹
Dir1.Refresh   ''刷新控件   ( 如果不这样弄的话,控件不会自动刷新 )
End Sub


Private Sub PutOnNotebook_Click()      ''''   ( Can doing)
On Error GoTo a:
Shell "notepad.exe " + File1.Path & "\" & File1.FileName, vbNormalFocus
Exit Sub

a:
MsgBox "没有记事本"
End Sub

Private Sub PutOntheProc_Click()      ''''   ( Can doing)
On Error Resume Next
Dim OpenProcPath As String
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName = "" Then Exit Sub
  OpenProcPath = CommonDialog1.FileName
Shell OpenProcPath & " " & File1.Path & "\" & File1.FileName, vbNormalFocus
End Sub

''''  #######################33    菜单 -> 对文件的操作
''###############      刷新控件

Private Sub NewDrive_Click()       ''''   ( Can doing)
Drive1.Refresh
End Sub

Private Sub NewBack_Click()      ''''   ( Can doing)
Dir1.Refresh
End Sub

Private Sub NewFile_Click()      ''''   ( Can doing)
File1.Refresh
End Sub

''###############      刷新控件

''''  #######################33    菜单 -> 对文件夹的操作

Private Sub Back_Click()       ''''   ( Can doing)
On Error Resume Next
Dir1.Path = Dir1.List(-2)   ''-1是本身,-2是上一级
End Sub

Private Sub OpenBack_Click()        ''''   ( Can doing)
Dir1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub DeleteBack_Click()      ''''   ( Can doing)
On Error GoTo a:
MsgBox Dir1.List(Dir1.ListIndex)
RmDir (Dir1.List(Dir1.ListIndex))    ''当文件夹有公文包时不能删除
Dir1.Refresh     ''刷新控件   ( 如果不这样弄的话,控件不会自动刷新 )
Exit Sub

a:
Dim NowDirectory As String
NowDirectory = CurDir
Reset
If RemoveDirectory(Dir1.List(Dir1.ListIndex)) = 0 Then MsgBox "删除失败"
ChDir (NowDirectory)
End Sub


''''  #######################33    菜单 -> 对文件夹的操作

Private Sub Scan_Click()       ''''   ( Can doing)
On Error Resume Next
Dim RootDisk As String
RootDisk = Drive1.Drive
Dim Directory As String
Directory = Dir1.Path
Reset

For i = 65 To 90
If PathFileExists(Chr(i) & ":\Autorun.inf") Then Warning Chr(i) & ":\Autorun.inf", Chr(i) & "\"
Next

MsgBox "程序扫描完毕,未发现存在的威胁", vbOKOnly, "OK"
Drive1.Drive = RootDisk
Dir1.Path = Directory
End Sub

Private Sub Abuot_Click()       ''''   ( Can doing)
MsgBox "U盘防火墙" + vbCrLf + "第" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + "版" + vbCrLf, vbOKOnly, "关于产品"
End Sub

Private Sub Look_Click()      ''''   ( Can doing)
Form2.Show
End Sub


''''  #######################33    菜单代码
''''  #########################

''''  #########################    文件控件代码
''''  #########################

Private Sub Dir1_Change()      ''''   ( Can doing)
ChDir (Dir1.Path)                  '''如果不设置这个东西的话,GetFileType就不会准确地运行
File1.Path = Dir1.Path
Dir1.ToolTipText = Dir1.Path
Label2.Caption = "该文件夹下共有:" & File1.ListCount & "个对象"   ''显示对象个数
Call FindAutorun    ''当文件夹位置改变时寻找是否有AutoRun.inf这个东西
End Sub

Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)        ''''   ( Can doing)
If Button = 2 Then
 If Dir1.List(-2) = "" Then
 Back.Enabled = False
 Else
 Back.Enabled = True
 End If
 
 If Dir1.ListIndex = -1 Then
 OpenBack.Enabled = False
 Else
 OpenBack.Enabled = True
 End If
 
Me.PopupMenu DoingBack    ''显示文件夹菜单
End If
End Sub

Private Sub Drive1_Change()       ''''   ( Can doing)         <--- Calling FindAutorun
On Error GoTo a:
Dir1.Path = Drive1.List(Drive1.ListIndex)
Exit Sub

a:
MsgBox "设备不可用"
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 Then File1.ToolTipText = File1.Path & File1.FileName & " -- " & FileType(File1.FileName, File1.Path, Drive1.Drive)
If File1.Selected(File1.ListIndex) = True Then     ''  如果在File1中已选中一项的话
If Button = 2 Then     ''如果按下的是右键
Me.PopupMenu Doing        ''显示文件菜单
End If
End If
End Sub


''''  #########################    文件控件代码
''''  #########################

Private Sub Form_Load()       ''''   ( Can doing)
Label2.Caption = "该文件夹下共有:" & File1.ListCount & "个对象"    ''显示对象个数
Dir1.ToolTipText = Dir1.Path
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)       ''''   ( Can doing)
Drive1.Refresh
Dir1.Refresh
File1.Refresh
End Sub


Private Sub Timer1_Timer()       ''''   ( Can doing)         <--- Calling DrivaType()
Label1.Caption = Dir1.Path   ''显示当前打开文件的路径
Label3.Caption = "该文件夹下共有:" & Dir1.ListCount & "个文件夹"   ''显示对象个数
Call DriveType(Drive1.Drive)
End Sub

Private Sub DriveType(ByVal RootPath As String)     '''' 返回根盘的类型      ( Can doing)
On Error Resume Next
Dim Drive As String
Drive = Trim(Left(RootPath, InStr(RootPath, "[") - 1))
If Drive = "" Then Drive = RootPath

Dim ReturnCode As Long
ReturnCode = GetDriveType(Drive)
If ReturnCode = 1 Then
Drive1.ToolTipText = RootPath + " 系统磁盘"
ElseIf ReturnCode = 2 Then
Drive1.ToolTipText = RootPath + " 可移动磁盘"
ElseIf ReturnCode = 3 Then
Drive1.ToolTipText = RootPath + " 本地磁盘"
ElseIf ReturnCode = 5 Then
Drive1.ToolTipText = RootPath + " CD-ROM"
ElseIf ReturnCode = 6 Then
Drive1.ToolTipText = RootPath + " RAM"
End If
End Sub

Private Function FileType(ByVal FileName As String, ByVal Path As String, ByVal Drive As String) As String '''' 返回文件的类型      ( Can doing)
ChDrive (Drive)
ChDir (Path)
Dim rCode As Long
Dim FileHwnd As Long
FileHwnd = lopen(FileName, 0)
rCode = GetFileType(FileHwnd)
If rCode = 1 Then
FileType = FileName + " 磁盘文件,返回代码:" & rCode
ElseIf rCode = 2 Then
FileType = FileName + " 控制台或打印机文件,返回代码:" & rCode
ElseIf rCode = 3 Then
FileType = FileName + " 管道文件,返回代码:" & rCode
Else
FileType = FileName + " 未知文件,返回代码:" & rCode
End If

lclose (FileHwnd)
End Function

Private Sub FindAutorun()       '''' 寻找File框中是否有 AutoRun -->利用循环语句  ( Can doing)
Dim i As Integer
For i = 0 To File1.ListCount - 1    ''最后那个是空项
Dim Name As String
Name = Trim(LCase(File1.List(i)))
Dim PathName, Path As String
PathName = Trim(File1.Path & "\" & File1.List(i))
Path = Trim(File1.Path)
If Name = "autorun.inf" Then
 Call Warning(PathName, Path)
End If
Next
End Sub

''############################################
''############################################   AutoRun 的文件内容查看引擎

Private Function AutoRunFindAnging(ByVal PathName As String) As String          ''''   ( Can doing)
On Error Resume Next
Dim ReturnCode As String
Dim Line As Long
Line = 0

Open PathName For Input As #1

Do While True
If EOF(1) = True Then Exit Do
Line = Line + 1
Dim Str As String
Line Input #1, Str
Dim Code As String
Code = CodeEqu(Str)
If Code <> "" Then ReturnCode = ReturnCode + CodeEqu(Str) + "   第: " + CStr(Line) + " 行" + vbCrLf

Loop

Close
AutoRunFindAnging = ReturnCode
End Function

Private Function CodeEqu(ByVal Str As String) As String           ''''   ( Can doing)
On Error Resume Next
Dim Back As String
Dim Value As String


Back = LCase(Trim(Left(Str, (InStrB(Str, "=") - 1) \ 2))) ''instr - 1 是让它不截取自身
Value = Trim(Mid(Str, InStr(Str, "=") + 1))   ''同上,不过表达式变了

''''   AutoRun .inf 的关键字眼
''''   该过程暂时提供回送关键信息
Dim AutoRun As String
AutoRun = "AutoRun 中含有  "
If Back = "" Or Value = "" Then Exit Function

If Back = "action" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "icon" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "label" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "open" Then

DeleteFile (Value)
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "useautoplay" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "shellexecute" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf Back = "shell" Then
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
ElseIf InStr(Back, "shell") > 0 Then

DeleteFile (Value)
AutoRun = AutoRun + Back + "  属性,指向 " + Value
CodeEqu = AutoRun
End If

End Function

''############################################   AutoRun 的文件内容查看引擎
''############################################

Private Sub Warning(ByVal PathName As String, ByVal Path As String)       ''''   ( Can doing)
On Error GoTo del:

Dim rCode As String
rCode = AutoRunFindAnging(PathName)
If rCode = "" Then
If (MsgBox("警告 ！" + vbCrLf + "在" + Path + "下含有AutoRun " + vbCrLf + "删除么", vbExclamation Or vbOKCancel, "检查到AutoRun") = vbOK) Then
DeleteFile (PathName)
File1.Refresh   ''刷新一下，不然的话用户会迷茫的
End If
Exit Sub

Else
 If (MsgBox("警告 ！" + vbCrLf + "在" + Path + "下含有AutoRun " + vbCrLf + "下面是AutoRun的内容" + vbCrLf + vbCrLf + rCode + vbCrLf + vbCrLf + "它可能会对您的电脑有危害" + vbCrLf + vbCrLf + "是否删除 ??", vbExclamation Or vbOKCancel, "检查到AutoRun", 0, 0) = vbOK) Then
  DeleteFile (PathName)    ''删除AutoRun
  File1.Refresh   ''刷新一下，不然的话用户会迷茫的
 End If
End If

Exit Sub

del:
MsgBox "删除出错"
File1.Refresh
End Sub

''可以循环提取File的文件
Private Sub ShowProc()      '''  实例过程-->可以利用循环快读取File中的数据
Dim i As Integer
Dim Proc As String
For i = 1 To File1.ListCount
Proc = Proc + File1.List(i) + Chr(10)
Next
Label2.ToolTipText = Proc
End Sub

