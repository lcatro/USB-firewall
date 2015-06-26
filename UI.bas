Attribute VB_Name = "UI"
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
cbSize As Long             ''  结构大小
hwnd As Long               ''  句柄
uId As Long                ''  结构ID
uFlags As Long             ''  传递给结构的标志
uCallbackMessage As Long   ''  回调信息
hIcon As Long              ''  Icon句柄(内存地址)
szTip As String * 64       ''  要显示在信息框内的信息
dwState As Long            ''  状态
dwStateMask As Long        ''  状态掩码
szInfo As String * 256     ''  气球提示框信息
uTimeout As Long           ''
uVerSion As Long           ''  版本
szInfoTitle As String * 64 ''  气球提示框标题
dwInfoFlags As Long        ''
guidItem As Long           ''  GUID选项
hBalloonIcon As Long       ''  气球图标的Icon句柄
End Type
''  Shell_NotifyIconA 函数dwMessage选参数
Private Const NIM_ADD = 0         ''  添加一个新状态栏
Private Const NIM_MODIFY = &H1    ''  修改状态栏的图标(包括它的结构信息)
Private Const NIM_DELETE = &H2    ''  删除一个存在的状态栏
Private Const NIM_SETFOCUS = &H3  ''  给存在的状态栏一个焦点
''  NOTIFYICONDATA结构信息Flag
Private Const NIF_MESSAGE = &H1   ''  开启回调信息(uCallbackMessage)
Private Const NIF_ICON = &H2      ''  开启ICON
Private Const NIF_TIP = &H4       ''  开启Tip信息
Private Const NIF_STATE = &H8     ''  dwStat,dwStateMask有效
Private Const NIF_INFO = &H10     ''  显示气球提示框
Private Const NIF_GUID = &H20     ''  高级IU选项[Win7或后面的版本可以用]
Private Const NIF_SHOWTIP = &H80  ''  立即显示Tip里面的信息
''  气球框将以下面的形式显示出来
Enum BalloonStyle
NIIF_NONE = 0         ''  没有任何风格
NIIF_INFO = &H1       ''  普通信息Icon提示
NIIF_WARNING = &H2    ''  警告的Icon提示
NIIF_ERROR = &H3      ''  错误的Icon提示
NIIF_USER = &H4       ''  使用用户自定义的Icon
NIIF_NOSOUND = &H10   ''  没有声音
End Enum

Dim ShellStruct As NOTIFYICONDATA

Sub CreateWindow(ByVal WindowHwnd As Long, ByVal IconHandle As Long)
ShellStruct.hwnd = WindowHwnd
ShellStruct.uFlags = NIF_ICON Or NIF_TIP
ShellStruct.hIcon = IconHandle
Shell_NotifyIconA NIM_ADD, ShellStruct
End Sub

Sub SetFocus()
Shell_NotifyIconA NIM_SETFOCUS, ShellStruct
End Sub

Sub ClearWindow()
Shell_NotifyIconA NIM_DELETE, ShellStruct
End Sub

Sub ChangeTip(ByVal TipString As String)
If Len(TipString) <= 64 Then
ShellStruct.szTip = TipString
Shell_NotifyIconA NIM_MODIFY, ShellStruct
End If
End Sub

Sub ShowTip()
ShellStruct.uFlags = NIF_SHOWTIP
Shell_NotifyIconA NIM_MODIFY, ShellStruct
End Sub
