[General]
SyntaxVersion=2
MacroID=1981fe98-b958-422b-a951-7ce591b11b9e
[Comment]

[Script]


// 获取游戏窗口句柄
Function 获取游戏窗口句柄
    // Hwnd0 不支持后台点击，但支持后台找图， Hwnd 不支持后台找图，但支持后台点击
    Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
    Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
    Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
    Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
    TracePrint Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd
    获取游戏窗口句柄 = Array(Hwnd0, Hwnd)
End Function

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
Sub ClickConfirmButton
    XY = findPicLocationTimes("common\确认.bmp" & "|" & "common\确认2.bmp", "确认", 10, 3000)
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
    If InStr(picPath, "|") = 0 Then 
    	picPath = kPicPath & picPath
        iCoord = Plugin.Bkgnd.FindPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.8)
    Else
        paths = Split(picPath, "|")
        For i = 0 To UBound(paths)
        	paths(i) = kPicPath & paths(i)
        Next
        picPath = Join(paths, "|")
        iCoord = Plugin.Bkgnd.FindMultiPic(Hwnd0,0, 0,1920, 1080,picPath, 0, 0.9)
    End If
    
    TracePrint picPath
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
    MsgBox message, styleNumber, kFileName
    EndScript
End Sub