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
MacroID=f205c1d4-266f-4451-beb4-02448937d386
Description=强者降临
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
/**
* @name: 火影OL 强者降临 脚本
* @description: 此脚本适用于《火影忍者ONLINE》游戏微端客户端，用于自动完成“强者降临”任务。
* @note: 运行此脚本需要在游戏首页界面；游戏窗口可后天运行，但不能最小化到任务栏；脚本运行时需保证处于第一窗口
* @version: 1.0.0
* @author: jason
* @contact: jason.wu.abcd@gmail.com
* @create: 2019-07-20
* @update: 2019-07-20
*/

[Script]
/**
* @name: 火影OL 强者降临 脚本
* @description: 此脚本适用于《火影忍者ONLINE》游戏微端客户端，用于自动完成“强者降临”任务。
* @note: 运行此脚本需要在游戏首页界面；游戏窗口可后天运行，但不能最小化到任务栏；脚本运行时需保证处于第一窗口
* @version: 1.0.0
* @author: jason
* @contact: jason.wu.abcd@gmail.com
* @create: 2019-07-20
* @update: 2019-07-20
*/

MsgBox "点击“确认”运行脚本", 4096
// 获取游戏窗口句柄
// Hwnd0 不支持后台点击，但支持后台找图， Hwnd 不支持后台找图，但支持后台点击
Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
TracePrint Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd

// Hwnd0 和 Hwnd 相同位置的Y坐标差
kOffsetY = 70
// 图片文件主目录
kPicPath = "D:\workspace\按键精灵脚本\火影OL脚本\"

enterBattleX = 0
enterBattleY = 0
fireX = 0
fireY = 0

For i=0 To 8
    // 1. 点击“强者降临”
    XY = findPicLocation(kPicPath & "强者降临\强者降临.bmp", "强者降临", True)
    Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+20, XY(1)+20)
    Delay 1500
    
    // 2. 判断奖励次数是否为0, 为0时结束脚本
    XY = findPicLocation(kPicPath & "强者降临\剩余次数为0.bmp", "强者降临", True)
    If IsNull(XY) = False Then 
        Call showEndMessage("剩余次数为0，脚本结束", False)
    End If

    // 3. 点击“阿修罗级”
    If i = 0 Then 
        XY = findPicLocation(kPicPath & "强者降临\阿修罗级_选中.bmp", "阿修罗级_选中", False)
        If IsNull(XY) Then 
            XY = findPicLocation(kPicPath & "强者降临\阿修罗级_未选中.bmp", "阿修罗级_未选中", True)
            Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1))
            Delay 1000
        End If
    End If
    
    // 4. 点击“进入战斗”
    If enterBattleX = 0 Then 
        Call findEnterBattleLocation()
    End If
    Call Plugin.Bkgnd.LeftClick(Hwnd, enterBattleX, enterBattleY)

    
    // 5. 点击“开战”
    Delay 6000
    Call ClickFireButton()
    Delay 3000
    
    // 6. 自动和2倍速
    If i = 0 Then 
        Call ClickAutoFireButton()
    End If
    Delay 110000
    
    // 7. 点击“确认”
    Call ClickConfirmButton()
    Delay 2000
    Call ClickConfirmButton()
    Delay 2000
Next

// 点击“进入战斗”
Sub findEnterBattleLocation()
    XY = findPicLocationTimes(kPicPath & "强者降临\进入战斗1.bmp" & "|" & kPicPath & "强者降临\进入战斗2.bmp", "进入战斗", 20, 100)
    enterBattleX = XY(0)
    enterBattleY = XY(1)
End Sub

// 点击“开战”
Sub ClickFireButton()
    If fireX = 0 Then 
        XY = findPicLocation(kPicPath & "common\开战.bmp", "开战", False)
        If IsNull(XY) = False Then 
            fireX = XY(0)
            fireY = XY(1)
        End If
    End If
    If fireX > 0 Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, fireX, fireY)
    Else
        Delay 3000
    End If
    Delay 3000
End Sub

// 点击自动战斗和2倍速
Sub ClickAutoFireButton()
    XY = findPicLocation(kPicPath & "common\自动.bmp", "自动", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 20)
        Delay 1000
    End If

    XY = findPicLocation(kPicPath & "common\速度.bmp", "速度", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+10, XY(1)+10)
    End If
End Sub

// 点击“确认”
Sub ClickConfirmButton()
    XY = findPicLocationTimes(kPicPath & "common\确认.bmp" & "|" & kPicPath & "common\确认2.bmp", "确认", 10, 3000)
    Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+20, XY(1)+10)
End Sub

// 在指定重复次数内查找指定图片，否则异常退出
Function findPicLocationTimes(picPath, name, time, delayTime)
    For time
        XY = findPicLocation(picPath, name, False)
        If IsNull(XY) Then 
            Delay delayTime
        Else 
            findPicLocationTimes = XY
            Exit Function
        End If
    Next
    Call showEndMessage("未发现“" & name & "”，脚本异常退出", True)
End Function

// 查找指定图片(可多个,"|"区分路径)坐标
Function findPicLocation(picPath, name, needEnd)
    Dim iCoord
    isMul = InStr(picPath, "|")
    If isMul = 0 Then 
        iCoord = Plugin.Bkgnd.FindPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.9)
    Else 
        iCoord = Plugin.Bkgnd.FindMultiPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.9)
    End If
    
    TracePrint name & "：" & iCoord
    XY = Split(iCoord, "|")
    If XY(0) > 0 and XY(1) > 0 Then 
        findPicLocation = Array(XY(0), XY(1)-kOffsetY)
    Else 
        If needEnd Then 
            Call showEndMessage("未找到“" & name & "”按钮", True)
        End If
        findPicLocation = Null
    End If
End Function

// 结束弹框提示
Sub showEndMessage(message, isFail)
    Dim styleNumber
    If isFail Then 
        styleNumber = 48+4096
    Else 
        styleNumber = 64+4096
    End If
    MsgBox message, styleNumber, kFileName
    EndScript
End Sub

