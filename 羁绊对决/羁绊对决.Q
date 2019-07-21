[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=f2b3b78d-5f65-4b84-b424-d3c56b21b018
Description=羁绊对决
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
/**
* @name: 火影OL 羁绊对决 脚本
* @description: 此脚本适用于《火影忍者ONLINE》游戏微端客户端，用于自动完成“羁绊对决”任务。
* @note: 运行此脚本需要在游戏首页界面；游戏窗口可后天运行，但不能最小化到任务栏；脚本运行时需保证处于第一窗口
* @version: 1.0.0
* @author: jason
* @contact: jason.wu.abcd@gmail.com
* @create: 2019-07-20
* @update: 2019-07-20
*/


MsgBox "点击“确认”运行脚本", 4096

// 获取游戏窗口句柄
Hwnds = Lib.HYOL.获取游戏窗口句柄
Hwnd0 = Hwnds(0)
Hwnd = Hwnds(1)

// Hwnd0 和 Hwnd 相同位置的Y坐标差
kOffsetY = 70
// 图片文件主目录
kPicPath = "D:\workspace\按键精灵脚本\火影OL脚本\"
kFileName = "羁绊对决"


// 1. 点击“限时活动”
XY = Lib.HYOL.findPicLocation("common\限时活动.bmp", "限时活动", True)
Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1)+20)
Delay 1500

// 2. 点击“羁绊对决”
For 3
    XY = Lib.HYOL.findPicLocation("羁绊对决\羁绊对决1.bmp" & "|" & "羁绊对决\羁绊对决2.bmp", "羁绊对决", False)
    If IsNull(XY) Then 
        down = Lib.HYOL.findPicLocation("common\限时活动_向下.bmp", "限时活动>向下箭头", True)
        Call Plugin.Bkgnd.LeftDown(Hwnd, down(0) + 10, down(1) + 10)
        Delay 100
        Call Plugin.Bkgnd.LeftUp(Hwnd, down(0) + 10, down(1) + 10)
    Else 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1))
        Exit For
    End If
Next
If IsNull(XY) Then 
    Call Lib.HYOL.showEndMessage("未找到“羁绊对决”, 异常退出", True)
End If
Delay 1000

// 3. 点击参加
XY = Lib.HYOL.findPicLocation("common\参加.bmp", "参加", True)
Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1))
Delay 5000

// 4. 点击“参加羁绊对决”
XY = Lib.HYOL.findPicLocationTimes("羁绊对决\参加羁绊对决.bmp", "参加羁绊对决", 5, 1000)
Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 40, XY(1) + 10)
Delay 3000

matchBattleBtnX = 0
matchBattleBtnY = 0

// 进入20次战斗
For 40
    // 5. 点击“匹配战斗”
    If matchBattleBtnX = 0 Then 
        XY = Lib.HYOL.findPicLocationTimes("羁绊对决\匹配战斗.bmp", "羁绊对决-匹配战斗", 5, 1000)
        matchBattleBtnX = XY(0)
        matchBattleBtnY = XY(1)
    End If
    If matchBattleBtnX > 0 Then 
    	Call Plugin.Bkgnd.LeftClick(Hwnd, matchBattleBtnX, matchBattleBtnY)
    End If
    Delay 2000
    
    // 点击自动
    Call Lib.HYOL.ClickAutoFireButton
    Delay 500
    
    // 6. 开启确认
    Call Lib.HYOL.ClickConfirmButton
    Delay 2000
    Call Lib.HYOL.ClickConfirmButton
    
    
    // 
Next





// 默认执行30分钟
//For 300
//    Call Plugin.Bkgnd.LeftClick(Hwnd, 852, 750)// 点击 "匹配战斗"
//    Delay 2000
//    Call Plugin.Bkgnd.LeftClick(Hwnd, 962, 590)// 点击 "点击确认"
//    Delay 2000
//    Call Plugin.Bkgnd.LeftClick(Hwnd, 962, 630)// 点击 "点击确认"
//    Delay 2000
//Next
