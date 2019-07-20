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
MacroID=9bd8c6fc-904d-4dfe-a45d-1a321f19da22
Description=组队副本
Enable=1
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=1
QUI=Form1
[Relative]
SetupOCXFile=
[Comment]
/**
* @name: 火影OL 组队副本脚本
* @description: 此脚本适用于《火影忍者ONLINE》游戏微端客户端，用于自动完成“组队副本”任务。
* @note: 运行此脚本需要在游戏首页界面，游戏窗口可后天运行，但不能最小化到任务栏
* @version: 1.0.0
* @author: jason
* @contact: jason.wu.abcd@gmail.com
* @create: 2019-07-20
* @update: 2019-07-20
*/

[Script]
/**
* @name: 火影OL 组队副本脚本
* @description: 此脚本适用于《火影忍者ONLINE》游戏微端客户端，用于自动完成“组队副本”任务。
* @note: 运行此脚本需要在游戏首页界面，游戏窗口可后天运行，但不能最小化到任务栏
* @version: 1.0.0
* @author: jason
* @contact: jason.wu.abcd@gmail.com
* @create: 2019-07-20
* @update: 2019-07-20
*/


// 获取游戏窗口句柄
// Hwnd0 不支持后台点击，但支持后台找图， Hwnd 不支持后台找图，但支持后台点击
Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
TracePrint Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd

// 页面模式选择框
//chooseMsg = MsgBox("选择窗口模式" & Chr(13) & Chr(10) & "‘是’：无头部窗口" & Chr(13) & Chr(10) & "‘否’：中等头部窗口" & Chr(13) & Chr(10) & "‘取消’：大头部窗口", 4096 + 67, "选择窗口模式")
// Hwnd0 和 Hwnd 相同位置的Y坐标差
KOffsetY = 70
//Select Case chooseMsg
//Case 6
//    KOffsetY = 70
//Case 7
//    KOffsetY = 70
//Case 2
//    KOffsetY = 70
//Case Else
//    MsgBox "选项值不合法：" & chooseMsg
//    EndScript
//End Select



// 1. 查找并点击“限时活动”
XY = findPicLocation("C:\Users\JASON\Desktop\限时活动.bmp", "限时活动", True)
Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1)+20)// 点击 "限时活动"
Delay 1500

// 2. 查找并点击“组队副本”
For 3
    XY = findMulPicLocation("C:\Users\JASON\Desktop\组队副本1.bmp|C:\Users\JASON\Desktop\组队副本2.bmp", "组队副本", False)
    If IsNull(XY) Then 
        down = findPicLocation("C:\Users\JASON\Desktop\限时活动_向下.bmp", "限时活动>向下箭头", True)
        Call Plugin.Bkgnd.LeftDown(Hwnd, down(0) + 10, down(1) + 10)
        Delay 100
        Call Plugin.Bkgnd.LeftUp(Hwnd, down(0) + 10, down(1) + 10)
        Delay 1000
    Else 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1))
        Exit For
    End If
Next
If IsNull(XY) Then 
    MessageBox "未找到“组队副本”"
    EndScript
End If

// 3. 查找并点击“参加”
Delay 1000
XY = findPicLocation("C:\Users\JASON\Desktop\组队副本参加.bmp", "组队副本>参加", True)
Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1))// 点击 "组队副本参加"
Delay 1500

fireX=0
fireY=0

For i = 0 To 13
    // 5. 点击困难本
    For 5
        XY = findPicLocation("C:\Users\JASON\Desktop\组队副本_困难.bmp", "组队副本>困难", False)
        If IsNull(XY) Then 
            Delay 2000
        Else 
            noTime = findPicLocation("C:\Users\JASON\Desktop\组队副本_剩余次数为0.bmp", "组队副本>困难", False)
            // 判断剩余次数是否为0， 为0时退出脚本
            If IsNull(noTime) = False Then 
                MsgBox "组队副本已完成", 64+4096, "火影脚本"
                EndScript
            End If
            Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0), XY(1)-100)
            Exit For
        End If
    Next
    
    // 6. 点击“不再提示”和“确认进入”
    Delay 1000
    If i = 0 Then 
        noTipLocation = findPicLocation("C:\Users\JASON\Desktop\组队副本_本次登录不再提示.bmp", "组队副本_本次登录不再提示", False)
        If IsNull(noTipLocation) = False Then 
            Call Plugin.Bkgnd.LeftClick(Hwnd, noTipLocation(0) + 10, noTipLocation(1) + 10)
            Delay 500
        End If
        confirmEnterLocation = findPicLocation("C:\Users\JASON\Desktop\组队副本_确认进入.bmp", "组队副本_确认进入", False)
        If IsNull(confirmEnterLocation) = False Then 
            Call Plugin.Bkgnd.LeftClick(Hwnd, confirmEnterLocation(0) + 30, confirmEnterLocation(1) + 20)
        End If
    End If
    
    // 7. 点击“开战”
    Delay 6000
    If fireX = 0 and fireY = 0 Then 
        Call findConfirmLocation()
    End If
    If fireX > 0 and fireY > 0 Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, fireX, fireY)
    Else 
        Delay 3000
    End If
    Delay 3000

    // 8. 点击“自动”和“2倍速”
    If i = 0 Then 
        XY = findPicLocation("C:\Users\JASON\Desktop\自动.bmp", "组队副本>自动", False)
        If IsNull(XY) = False Then 
            Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 20)
            Delay 1000
        End If

        XY = findPicLocation("C:\Users\JASON\Desktop\速度.bmp", "组队副本>速度", False)
        If IsNull(XY) = False Then 
            Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+10, XY(1)+10)
        End If
    End If
    Delay 30000
    
    // 9. 点击“返回游戏”
    For 10
        XY = findPicLocation("C:\Users\JASON\Desktop\组队副本_返回游戏.bmp", "组队副本>返回游戏", False)
        If IsNull(XY) Then 
            Delay 3000
        Else 
            Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 10)
            Exit For
        End If
    Next
    
    Delay 5000
Next

// 查找开战按钮位置
Sub findConfirmLocation()
    XY = findPicLocation("C:\Users\JASON\Desktop\组队副本_开战.bmp", "组队副本>开战", False)
    If IsNull(XY) = False Then 
        fireX = XY(0) + 20
        fireY = XY(1) + 20
    End If
End Sub

Function findPicLocation(picPath, name, needEnd)
    iCoord = Plugin.Bkgnd.FindPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.7)
    TracePrint name & "：" & iCoord
    XY = Split(iCoord, "|")
    If XY(0) > 0 and XY(1) > 0 Then 
        findPicLocation = Array(XY(0), XY(1)-KOffsetY)
    Else 
        If needEnd Then 
            MessageBox "未找到“" & name & "”按钮"
            EndScript
        End If
        findPicLocation = Null
    End If
End Function

Function findMulPicLocation(picPath, name, needEnd)
    iCoord = Plugin.Bkgnd.FindMultiPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.7)
    TracePrint name & "：" & iCoord
    XY = Split(iCoord, "|")
    If XY(0) > 0 and XY(1) > 0 Then 
        findMulPicLocation = Array(XY(0), XY(1)-KOffsetY)
    Else 
        If needEnd Then 
            MessageBox "未找到“" & name & "”按钮"
            EndScript
        End If
        findMulPicLocation = Null
    End If
End Function
