[General]
SyntaxVersion=2
MacroID=1981fe98-b958-422b-a951-7ce591b11b9e
[Comment]

[Script]


// ��ȡ��Ϸ���ھ��
Function ��ȡ��Ϸ���ھ��
    // Hwnd0 ��֧�ֺ�̨�������֧�ֺ�̨��ͼ�� Hwnd ��֧�ֺ�̨��ͼ����֧�ֺ�̨���
    Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
    Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
    Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
    Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
    TracePrint Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd
    ��ȡ��Ϸ���ھ�� = Array(Hwnd0, Hwnd)
End Function

// ����Զ�ս����2����
Sub ClickAutoFireButton()
    XY = findPicLocation(kPicPath & "common\�Զ�.bmp", "�Զ�", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 20)
        Delay 1000
    End If

    XY = findPicLocation(kPicPath & "common\�ٶ�.bmp", "�ٶ�", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+10, XY(1)+10)
    End If
End Sub

// �����ȷ�ϡ�
Sub ClickConfirmButton
    XY = findPicLocationTimes("common\ȷ��.bmp" & "|" & "common\ȷ��2.bmp", "ȷ��", 10, 3000)
    Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+20, XY(1)+10)
End Sub

// ��ָ���ظ������ڲ���ָ��ͼƬ�������쳣�˳�
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
    Call showEndMessage("δ���֡�" & name & "�����ű��쳣�˳�", True)
End Function

// ����ָ��ͼƬ(�ɶ��,"|"����·��)����
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
    TracePrint name & "��" & iCoord
    XY = Split(iCoord, "|")
    If XY(0) > 0 and XY(1) > 0 Then 
        findPicLocation = Array(XY(0), XY(1)-kOffsetY)
    Else 
        If needEnd Then 
            Call showEndMessage("δ�ҵ���" & name & "����ť", True)
        End If
        findPicLocation = Null
    End If
End Function

// ����������ʾ
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