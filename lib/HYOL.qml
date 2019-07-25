[General]
SyntaxVersion=2
MacroID=1981fe98-b958-422b-a951-7ce591b11b9e
[Comment]

[Script]


/*
----------------------------------------��ʼ��----------------------------------------
*/
// ����ȫ�ֳ���
Sub SetGlobalConstant
    Const a ="aaa"
    MsgBox "�����ȷ�ϡ����нű�", 4096
    // Hwnd0 �� Hwnd ��ͬλ�õ�Y�����
    kOffsetY = 70
    // ͼƬ�ļ���Ŀ¼
    kPicPath = "C:\Users\JASON-BOOK\Documents\workspace\��ӰOL�ű�\HYOL\"
    // ʹ��Dim��Const���������(����)�󣬱�����������ֻ�ں������ڲ���Ч��������Dim���壬����������Ҳ�ܷ��ʵ�
    // �Ƿ�Ϊ���߼�ģʽ��
    Dim isAdvance
    isAdvance = False
    // ��Ļ�ֱ���
    kScreenW = 1400
    kScreenH = 800
    
    // ��ȡ��Ϸ���ھ��
    Dim Hwnd1, Hwnd2
    // Hwnd0 ��֧�ֺ�̨�������֧�ֺ�̨��ͼ�� Hwnd ��֧�ֺ�̨��ͼ����֧�ֺ�̨���
    Hwnd0 = Plugin.Window.Find("MainView_9F956014-12FC-42d8-80C7-9A90D4D567E3", 0)
    If isAdvance Then
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "CefBrowserWindow", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "Chrome_WidgetWin_0", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd2, 0, "Chrome_RenderWidgetHostHWND", 0)
        TracePrint "�߼�ģʽ" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd
    Else 
        kOffsetY = kOffsetY + 50
        Dim Hwnd3, Hwnd4, Hwnd5
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "iecontainerclass", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "AtlAxWin120", 0)
        Hwnd3 = Plugin.Window.FindEx(Hwnd2, 0, "Shell Embedding", 0)
        Hwnd4 = Plugin.Window.FindEx(Hwnd3, 0, "Shell DocObject View", 0)
        Hwnd5 = Plugin.Window.FindEx(Hwnd4, 0, "Internet Explorer_Server", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd5, 0, "MacromediaFlashPlayerActiveX", 0)
        TracePrint "����ģʽ��" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd3 & " " & Hwnd4 & " " & Hwnd5 & " " & Hwnd
    End If
End Sub


/*
-------------------------------����ģ��---------------------------------
*/
 
// ����������
Sub HYLeftClick(X, Y)
    Call Plugin.Bkgnd.LeftClick(Hwnd, X, Y)
End Sub

// ����������
Sub HYLeftDown(X, Y)
    Call Plugin.Bkgnd.LeftDown(Hwnd, X, Y)
End Sub

// ����������
Sub HYLeftUp(X, Y)
    Call Plugin.Bkgnd.LeftUp(Hwnd, X , Y)
End Sub

// ����ƶ�
Sub HYMoveTo(X, Y)
	Call Plugin.Bkgnd.MoveTo(Hwnd, X, Y)
End Sub

// ��ָ���ظ������ڲ���ָ��ͼƬ�������쳣�˳�
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
    	Call showEndMessage("δ���֡�" & name & "�����ű��쳣�˳�", True)
    Else
    	findPicLocationTimes = Null
  End If
End Function

// ����ָ��ͼƬ(�ɶ��,"|"����·��)����
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
    MsgBox message, styleNumber, "��ӰOL�ű�"
    ExitScript
End Sub

/*
--------------------------------ҵ��ģ��--------------------------------------
*/

// �������ս��
Sub ClickFireButton
    Dim XY
    XY = findPicLocationTimes("common\��ս.bmp","��ս", 5,1000, False)
    If IsNull(XY) = False Then 
        Call Lib.HYOL.HYLeftClick(XY(0)+40, XY(1)+15)
    Else 
        // ���������������� �������ִ�к������
        Delay 3000
    End If
End Sub

// ����Զ�ս����2����
Sub ClickAutoFireButton()
    Dim XY
    XY = findPicLocation("common\�Զ�.bmp", "�Զ�", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0) + 20, XY(1) + 20)
        Delay 1000
    End If
    XY = findPicLocation("common\�ٶ�.bmp", "�ٶ�", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+10, XY(1)+10)
    End If
End Sub

// �����ȷ�ϡ�
Sub ClickConfirmButton
    Dim XY
    XY = findPicLocation("common\ȷ��.bmp" & "|" & "common\ȷ��2.bmp" & "|" & "common\ȷ��.bmp", "ȷ��", False)
    If IsNull(XY) = False Then 
        Call Plugin.Bkgnd.LeftClick(Hwnd, XY(0)+20, XY(1)+10)
    End If
End Sub

// �����ʱ���ָ��ѡ�������μ�
Sub ClickTimeActivitiesOption(picPaths, name)
    Dim XY, i
    // �����ʱ�
    XY = findPicLocation("common\��ʱ�.bmp", "��ʱ�", True)
    Call HYLeftClick(XY(0), XY(1)+20)
    
    // ����ƶ�ѡ��
    Delay 1500
    Const count = 3
    For i = 0 To count
        XY = findPicLocation(picPaths, name, False)
        If IsNull(XY) Then 
            down = findPicLocation("common\��ʱ�_����.bmp", "��ʱ�>���¼�ͷ", True)
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
        Call showEndMessage("δ�ҵ����Ծ���, �쳣�˳�", True)
    End If
    
    // ����μ�
    Delay 1000
    XY = findPicLocation("common\�μ�.bmp", "�μ�", True)
    Call HYLeftClick(XY(0), XY(1))
End Sub
