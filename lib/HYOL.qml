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
    Const a ="aaa"
    MsgBox "点击“确认”运行脚本", 4096
    // Hwnd0 和 Hwnd 相同位置的Y坐标差
    kOffsetY = 70
    // 图片文件主目录
    kPicPath = "C:\Users\JASON-BOOK\Documents\workspace\火影OL脚本\HYOL\"
    // 使用Dim（Const）定义变量(常量)后，变量（常量）只在函数体内部有效；若不用Dim定义，则函数体外面也能访问到
    // 是否为“高级模式”
    Dim isAdvance
    isAdvance = False
    // 屏幕分辨率
    kScreenW = 1400
    kScreenH = 800
    
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
        kOffsetY = kOffsetY + 50
        Dim Hwnd3, Hwnd4, Hwnd5
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "iecontainerclass", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "AtlAxWin120", 0)
        Hwnd3 = Plugin.Window.FindEx(Hwnd2, 0, "Shell Embedding", 0)
        Hwnd4 = Plugin.Window.FindEx(Hwnd3, 0, "Shell DocObject View", 0)
        Hwnd5 = Plugin.Window.FindEx(Hwnd4, 0, "Internet Explorer_Server", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd5, 0, "MacromediaFlashPlayerActiveX", 0)
        TracePrint "基础模式：" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd3 & " " & Hwnd4 & " " & Hwnd5 & " " & Hwnd
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
        iCoord = Plugin.Bkgnd.FindPic(Hwnd0,0, 0,kScreenW, kScreenH, picPath, 0, 0.6)
    Else 
        Dim paths, i
        paths = Split(picPath, "|")
        For i = 0 To UBound(paths)
            paths(i) = kPicPath & paths(i)
        Next
        picPath = Join(paths, "|")
        iCoord = Plugin.Bkgnd.FindMultiPic(Hwnd0,0, 0,kScreenW, kScreenH, picPath, 0, 0.6)
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
--------------------------------业务模块--------------------------------------
*/

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

// 点击“确认”
Sub ClickConfirmButton
    Dim XY
    XY = findPicLocation("common\确认.bmp" & "|" & "common\确认2.bmp" & "|" & "common\确定.bmp", "确认", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+20, XY(1)+10)
    End If
End Sub

// 点击限时活动中指定选项，并点击参加
Sub ClickTimeActivitiesOption(picPaths, name)
    Dim XY, i
    // 点击限时活动
    XY = findPicLocation("common\限时活动.bmp", "限时活动", True)
    Call HYLeftClick(XY(0), XY(1)+20)
    
    // 点击制定选项
    Delay 1500
    Const count = 3
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
