[General]
SyntaxVersion=2
MacroID=1981fe98-b958-422b-a951-7ce591b11b9e
[Comment]

[Script]


/*
----------------------------------------初始化----------------------------------------
*/
// 设置全局常量
Sub SetGlobalConstant
	kHeader1H = 70
	kHeader2H = 50
	
    MsgBox "点击“确认”运行脚本", 4096
    // Hwnd0 和 Hwnd 相同位置的Y坐标差
    kOffsetY = kHeader1H
    // 图片文件主目录
    kPicPath = "C:\Users\JASON-BOOK\Documents\workspace\火影OL脚本\HYOL\"
    // 使用Dim（Const）定义变量(常量)后，变量（常量）只在函数体内部有效；若不用Dim定义，则函数体外面也能访问到
    // 是否为“高级模式”
    Dim isAdvance
    isAdvance = False
    // 屏幕分辨率
    kScreenW = 1400
    kScreenH = 800
    
    Hwnd_2 = 0
    // 获取游戏窗口句柄
    Dim Hwnd1, Hwnd2
    // Hwnd0 不支持后台点击，但支持后台找图， Hwnd 不支持后台找图，但支持后台点击
    Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
    If isAdvance Then
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
        TracePrint "高级模式" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd
    Else 
        Dim Hwnd3, Hwnd4, Hwnd5
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "iecontainerclass", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "AtlAxWin120", 0)
        Hwnd3 = Plugin.Window.FindEx(Hwnd2, 0, "Shell Embedding", 0)
        Hwnd4 = Plugin.Window.FindEx(Hwnd3, 0, "Shell DocObject View", 0)
        Hwnd5 = Plugin.Window.FindEx(Hwnd4, 0, "Internet Explorer_Server", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd5, 0, "MacromediaFlashPlayerActiveX", 0)
        TracePrint "基础模式：" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd3 & " " & Hwnd4 & " " & Hwnd5 & " " & Hwnd
        
        Color = Plugin.Bkgnd.GetPixelColor(Hwnd0, 13, 84)
		If Color = "404040" Then 
			TracePrint "有header2"
			kOffsetY = kOffsetY + kHeader2H
		End If
    End If
End Sub


/*
-------------------------------基础模块---------------------------------
*/
 
// 鼠标左键单击
Sub HYLeftClick(X, Y)
    Call Plugin.Bkgnd.LeftClick(Hwnd, X, Y)
End Sub

// 鼠标左键按下
Sub HYLeftDown(X, Y)
    Call Plugin.Bkgnd.LeftDown(Hwnd, X, Y)
End Sub

// 鼠标左键弹起
Sub HYLeftUp(X, Y)
    Call Plugin.Bkgnd.LeftUp(Hwnd, X , Y)
End Sub

// 鼠标移动
Sub HYMoveTo(X, Y)
    Call Plugin.Bkgnd.MoveTo(Hwnd, X, Y)
End Sub

// 在指定重复次数内查找指定图片，否则异常退出
Function findPicLocationTimes(picPath, name, time, delayTime, needEnd)
    Dim XY
    For time
        XY = findPicLocation(picPath, name, False)
        If IsNull(XY) Then 
            Delay delayTime
        Else 
            findPicLocationTimes = XY
            Exit Function
        End If
    Next
    If needEnd = True Then 
        Call showEndMessage("未发现“" & name & "”，脚本异常退出", True)
    Else
        findPicLocationTimes = Null
    End If
End Function

// 查找指定图片(可多个,"|"区分路径)坐标
Function findPicLocation(picPath, name, needEnd)
    Dim iCoord, XY
    If InStr(picPath, "|") = 0 Then 
        picPath = kPicPath & picPath
        iCoord = Plugin.Bkgnd.FindPic(Hwnd0,0, 0,kScreenW, kScreenH, picPath, 0, 0.8)
    Else 
        Dim paths, i
        paths = Split(picPath, "|")
        For i = 0 To UBound(paths)
            paths(i) = kPicPath & paths(i)
        Next
        picPath = Join(paths, "|")
        iCoord = Plugin.Bkgnd.FindMultiPic(Hwnd0,0, 0,kScreenW, kScreenH, picPath, 0, 0.8)
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
        styleNumber = 16+4096
    Else 
        styleNumber = 64+4096
    End If
    MsgBox message, styleNumber, "火影OL脚本"
    ExitScript
End Sub

/*
-------------------------------- 通用业务模块 --------------------------------------
*/

// 切换游戏窗口
Sub SwitchGameWindow(index)
	Call Plugin.Bkgnd.MoveTo(Hwnd0, 125*index+40, 25) 
	Call Plugin.Bkgnd.LeftClick(Hwnd0, 125 * index + 40, 25)
	// 重新获取窗口句柄还是第一个窗口的，无法获取第二个窗口的句柄
End Sub

// 点击“开战”
Sub ClickFireButton
    Dim XY
    XY = findPicLocationTimes("common\开战.bmp","开战", 5,1000, False)
    If IsNull(XY) = False Then 
        Call Lib.HYOL.HYLeftClick(XY(0)+40, XY(1)+15)
    Else 
        // 针对网络慢的情况， 多等三秒执行后面操作
        Delay 3000
    End If
End Sub

// 点击自动战斗和2倍速
Sub ClickAutoFireButton()
	Dim XY
    XY = findPicLocation("common\自动.bmp", "自动", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 20)
        Delay 1000
    End If
    
    XY = findPicLocation("common\速度.bmp", "速度", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+10, XY(1)+10)
    End If
End Sub

Sub ClickAutoFireButtonInAllWindow
	Call ClickAutoFireButton()
	Delay 1000
	Call SwitchGameWindow(2)
	Delay 2000
	
	Dim XY
	XY = findPicLocation("common\自动.bmp", "自动", False)
	If IsNull(XY) = False Then 
		MoveTo XY(0)+10, XY(1)+kOffsetY+10
		LeftClick 1
	End If
	
	Delay 1000
	Call SwitchGameWindow(1)
End Sub

// 点击限时活动中指定选项，并点击参加
Sub ClickTimeActivitiesOption(picPaths, name)
    Dim XY, i
    // 点击限时活动
    XY = findPicLocation("common\限时活动.bmp", "限时活动", True)
    Call HYLeftClick(XY(0), XY(1)+20)
    
    // 点击制定选项
    Delay 1500
    Dim count 
    count = 3
    For i = 0 To count
        XY = findPicLocation(picPaths, name, False)
        If IsNull(XY) Then 
            down = findPicLocation("common\限时活动_向下.bmp", "限时活动>向下箭头", True)
            Call HYLeftDown(down(0) + 10, down(1) + 10)
            Delay 100
            Call HYLeftUp(down(0) + 10, down(1) + 10)
            If Not (i = count)  Then 
                Delay 1000
            End If
        Else 
            Call HYLeftClick(XY(0), XY(1))
            Exit For
        End If
    Next
    If IsNull(XY) Then
        Call showEndMessage("未找到“羁绊对决”, 异常退出", True)
    End If
    
    // 点击参加
    Delay 1000
    XY = findPicLocation("common\参加.bmp", "参加", True)
    Call HYLeftClick(XY(0), XY(1))
End Sub

/*
--------------------------------------------- 组队 -----------------------------------------------
*/

Sub 组队本(isEnd)
    // 1. 限时活动中参加“组队副本”
    Call Lib.HYOL.ClickTimeActivitiesOption("组队副本\组队副本1.bmp" & "|" & "组队副本\组队副本2.bmp", "组队副本")
    Delay 2000
	
    Dim 组队副本困难X, 组队副本困难Y, XY, noTime
    组队副本困难X = 0
    组队副本困难Y = 0 

    For i = 0 To 13
        // 判断剩余次数是否为0， 为0时退出脚本
        noTime = Lib.HYOL.findPicLocation("组队副本\组队副本_剩余次数为0.bmp", "组队副本>剩余次数为0", False)
        If IsNull(noTime) = False Then 
            If isEnd Then 
                Call Lib.HYOL.showEndMessage("剩余次数为0，组队副本已完成", False)
            Else 
                Exit Sub
            End If
            
        End If
    
        // 点击困难本, 40秒未发现则异常退出
        If 组队副本困难X = 0 Then 
            XY = Lib.HYOL.findPicLocationTimes("组队副本\组队副本_困难.bmp", "组队副本>困难", 10, 2000, True)
            组队副本困难X = XY(0)
            组队副本困难Y = XY(1) - 100
        End If
        Call Lib.HYOL.HYLeftClick(组队副本困难X, 组队副本困难Y)
    
        // 点击“不再提示”和“确认进入”
        If i = 0 Then 
            Delay 1000
            Dim noTipLocation, confirmEnterLocation
            noTipLocation = Lib.HYOL.findPicLocation("组队副本\组队副本_本次登录不再提示.bmp", "组队副本>本次登录不再提示", False)
            If IsNull(noTipLocation) = False Then 
                Call Lib.HYOL.HYLeftClick(noTipLocation(0) + 10, noTipLocation(1) + 10)
                Delay 500
            End If
            confirmEnterLocation = Lib.HYOL.findPicLocation("组队副本\组队副本_确认进入.bmp", "组队副本>确认进入", False)
            If IsNull(confirmEnterLocation) = False Then 
                Call Lib.HYOL.HYLeftClick(confirmEnterLocation(0) + 30, confirmEnterLocation(1) + 20)
            End If
        End If
    
        // 点击“开战”
        Delay 6000
        Call Lib.HYOL.ClickFireButton
        Delay 3000

        // 点击“自动”和“2倍速”
        If i = 0 Then 
            Call Lib.HYOL.ClickAutoFireButtonInAllWindow
            
        End If
    
        // 预计战斗时间
        Delay 30000
    
        // 点击“返回游戏”
        XY = Lib.HYOL.findPicLocationTimes("组队副本\返回游戏.bmp", "组队副本>返回游戏", 30, 3000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 50, XY(1) + 20)
        Delay 4000
    Next
End Sub

/*
----------------------------------------- 生存演戏 --------------------------------------------
*/
Sub 生存演戏(isEnd)
    Dim XY
    // 点击进入 生存演戏
    XY = Lib.HYOL.findPicLocation("生存演戏\生存演戏.bmp", "生存演戏", True)
    Call Lib.HYOL.HYLeftClick(XY(0)+20, XY(1)+20)
    Delay 3000

    // 判断是否已完成
    XY = Lib.HYOL.findPicLocation("生存演戏\已完成.bmp", "生存演戏-已完成", False)
    If IsNull(XY) = False Then 
        Goto 检测是否还有次数
    End If

    Rem 开始

    // 判断是否已经选中了“极限”
    XY = Lib.HYOL.findPicLocation("生存演戏\极限_选中.bmp", "生存演戏-极限_选中", False)
    If IsNull(XY) Then 
        // 判断 极限 是否已激活
        XY = Lib.HYOL.findPicLocation("生存演戏\极限_激活.bmp", "生存演戏-极限_激活", False)
        If IsNull(XY) Then 
            XY = Lib.HYOL.findPicLocation("生存演戏\极限_锁定.bmp", "生存演戏-极限_锁定", True)
            Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
            Delay 1000
            XY = Lib.HYOL.findPicLocation("生存演戏\是.bmp", "生存演戏_是", True)
            Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
            Delay 2000
            XY = Lib.HYOL.findPicLocation("生存演戏\极限_激活.bmp", "生存演戏-极限_激活", True)
        End If
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
        Delay 1000
        XY = Lib.HYOL.findPicLocation("生存演戏\是.bmp", "生存演戏_是", True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
        Delay 2000
    End If

    Dim picName
    For i = 1 To 4
        If i = 1 Then 
            picName = "SS本"
        ElseIf i = 2 or i = 3 Then
            picName = "S本"
        Else 
            picName = "A本"
        End If
        // 点击 SS
        XY = Lib.HYOL.findPicLocation("生存演戏\" & picName & ".bmp", "生存演戏_" & picName, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
        Delay 10000

        // 点击 开战
        XY = Lib.HYOL.findPicLocation("生存演戏\开战.bmp", "生存演戏_开战", True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 5000
	
        // 点击 自动 和 2倍速
        If i = 1 Then 
            Call Lib.HYOL.ClickAutoFireButton
        End If
        // 战斗时间最短20秒
        Delay 20000
    
        // 检测 确定 按钮
        XY = Lib.HYOL.findPicLocationTimes("生存演戏\确定.bmp", "生存演戏-确定", 30, 3000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 2000
    Next
    Delay 3000
    // 点击一键领取
    XY = Lib.HYOL.findPicLocationTimes("生存演戏\一键领取.bmp", "生存演戏-一键领取", 3, 2000, True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 15)
    Delay 1000
    XY = Lib.HYOL.findPicLocation("生存演戏\是.bmp", "生存演戏-是", True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
    Delay 3000
    XY = Lib.HYOL.findPicLocation("common\确定2.bmp", "生存演戏-确定2", True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
    Delay 1000

    Rem 检测是否还有次数
    // 判断是否还有次数
    XY = Lib.HYOL.findPicLocation("生存演戏\重置.bmp", "生存演戏-重置", False)
    If IsNull(XY) = False Then 
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 1000
        Goto 开始
    End If
	
    If isEnd Then 
        Call Lib.HYOL.showEndMessage("生存演戏完成", False)
    End If

End Sub

/*
------------------------------------------------ 羁绊对决 ----------------------------------------------------
*/

Sub 羁绊对决(isEnd)
    // 1. 限时活动中参加“羁绊对决”
    Call Lib.HYOL.ClickTimeActivitiesOption("羁绊对决\羁绊对决1.bmp" & "|" & "羁绊对决\羁绊对决2.bmp", "羁绊对决")

    Delay 5000
	
    Dim XY
    // 2. 点击“参加羁绊对决”
    XY = Lib.HYOL.findPicLocationTimes("羁绊对决\参加羁绊对决.bmp", "参加羁绊对决", 5, 1000, True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 10)
    Delay 4000
	
    Dim matchBattleBtnX, matchBattleBtnY, 确认撤退X, 确认撤退Y, 确定X, 确定Y
    matchBattleBtnX = 0
    matchBattleBtnY = 0
    确认撤退X = 0 
    确认撤退Y = 0
    确定X = 0
    确定Y = 0


    // 胜利次数小于12次, 人头数小于96
    Dim i, 匹配到战斗
    i=0
    While i<12
        // 3. 点击“匹配战斗”
        If matchBattleBtnX = 0 Then 
            XY = Lib.HYOL.findPicLocationTimes("羁绊对决\匹配战斗.bmp", "羁绊对决-匹配战斗",5,2000,True)
            matchBattleBtnX = XY(0)
            matchBattleBtnY = XY(1)
        End If
        Call Lib.HYOL.HYLeftClick(matchBattleBtnX, matchBattleBtnY)
        Delay 2000
    
        // 判断是否匹配到战斗
        Rem 查询开始
        匹配到战斗 = False
        For 10
            // 判断是否匹配失败
            XY = Lib.HYOL.findPicLocation("羁绊对决\匹配战斗.bmp", "羁绊对决-匹配战斗",False)
            If IsNull(XY) = False Then 
                Call Lib.HYOL.HYLeftClick(matchBattleBtnX, matchBattleBtnY)
                Delay 2000
                Goto 查询开始
            End If
            Delay 100
            // 判断对方是否已撤退
            XY = Lib.HYOL.findPicLocation("common\确认.bmp", "羁绊对决-确认", False)
            If IsNull(XY) = False Then 
                匹配到战斗 = True
                Exit For
            End If
		
            // 判断是否进入了战斗
            XY = Lib.HYOL.findPicLocation("common\速度_灰色.bmp", "羁绊对决>速度灰色", False)
            If IsNull(XY) Then 
                Delay 10000
            Else 
                匹配到战斗 = True
                Exit For
            End If
        Next
        If 匹配到战斗 = False Then 
            Call Lib.HYOL.showEndMessage("2分钟内未匹配到战斗，脚本退出", True)
        End If
        Delay 1000
    
        // 点击自动
        If i = 0 or i = 1 Then 
            Call Lib.HYOL.ClickAutoFireButtonInAllWindow
            Delay 1000
        End If
    
        // 判断是否有撤退按钮，有的话就直接撤退
        XY = Lib.HYOL.findPicLocation("common\撤退.bmp", "羁绊对决-撤退", False)
        If IsNull(XY) Then 
            i = i + 1
            // 延时10秒，判断对方是否有撤退
            Delay 10000
            XY = Lib.HYOL.findPicLocation("common\确认.bmp", "羁绊对决-确认", False)
            If IsNull(XY) Then 
                Delay 50000
            End If
        Else 
            // 点击撤退
            Call Lib.HYOL.HYLeftClick(XY(0), XY(1))
            Delay 1000
            If 确认撤退X = 0 Then 
                XY = Lib.HYOL.findPicLocation("common\确认撤退.bmp", "羁绊对决-确认撤退", True)
                确认撤退X = XY(0)+20
                确认撤退Y = XY(1)+10
            End If
            Call Lib.HYOL.HYLeftClick(确认撤退X, 确认撤退Y)
            Delay 3000
        End If
    
        // 点击“确认”按钮
        XY = Lib.HYOL.findPicLocationTimes("common\确认.bmp", "羁绊对决-确认", 100, 5000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 10, XY(1) + 5)
        Delay 5000
    
        // 点击"确定"按钮
        If 确定X = 0 Then 
            XY = Lib.HYOL.findPicLocation("common\确定2.bmp|common\确定.bmp", "羁绊对决-确定", True)
            确定X = XY(0) + 10
            确定Y = XY(1) + 5
        End If
        Call Lib.HYOL.HYLeftClick(确定X, 确定Y)
        Delay 2000
    Wend
    If isEnd Then 
        Call Lib.HYOL.showEndMessage("羁绊对决结束", False)
    End If
End Sub

