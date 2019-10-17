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
	kHeader1H = 70
	kHeader2H = 50
	
    MsgBox "�����ȷ�ϡ����нű�", 4096
    // Hwnd0 �� Hwnd ��ͬλ�õ�Y�����
    kOffsetY = kHeader1H
    // ͼƬ�ļ���Ŀ¼
    kPicPath = "C:\Users\JASON-BOOK\Documents\workspace\��ӰOL�ű�\HYOL\"
    // ʹ��Dim��Const���������(����)�󣬱�����������ֻ�ں������ڲ���Ч��������Dim���壬����������Ҳ�ܷ��ʵ�
    // �Ƿ�Ϊ���߼�ģʽ��
    Dim isAdvance
    isAdvance = False
    // ��Ļ�ֱ���
    kScreenW = 1400
    kScreenH = 800
    
    Hwnd_2 = 0
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
        Dim Hwnd3, Hwnd4, Hwnd5
        Hwnd1 = Plugin.Window.FindEx(Hwnd0, 0, "iecontainerclass", 0)
        Hwnd2 = Plugin.Window.FindEx(Hwnd1, 0, "AtlAxWin120", 0)
        Hwnd3 = Plugin.Window.FindEx(Hwnd2, 0, "Shell Embedding", 0)
        Hwnd4 = Plugin.Window.FindEx(Hwnd3, 0, "Shell DocObject View", 0)
        Hwnd5 = Plugin.Window.FindEx(Hwnd4, 0, "Internet Explorer_Server", 0)
        Hwnd = Plugin.Window.FindEx(Hwnd5, 0, "MacromediaFlashPlayerActiveX", 0)
        TracePrint "����ģʽ��" & Hwnd0 & " " & Hwnd1 & " " & Hwnd2 & " " & Hwnd3 & " " & Hwnd4 & " " & Hwnd5 & " " & Hwnd
        
        Color = Plugin.Bkgnd.GetPixelColor(Hwnd0, 13, 84)
		If Color = "404040" Then 
			TracePrint "��header2"
			kOffsetY = kOffsetY + kHeader2H
		End If
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
-------------------------------- ͨ��ҵ��ģ�� --------------------------------------
*/

// �л���Ϸ����
Sub SwitchGameWindow(index)
	Call Plugin.Bkgnd.MoveTo(Hwnd0, 125*index+40, 25) 
	Call Plugin.Bkgnd.LeftClick(Hwnd0, 125 * index + 40, 25)
	// ���»�ȡ���ھ�����ǵ�һ�����ڵģ��޷���ȡ�ڶ������ڵľ��
End Sub

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

Sub ClickAutoFireButtonInAllWindow
	Call ClickAutoFireButton()
	Delay 1000
	Call SwitchGameWindow(2)
	Delay 2000
	
	Dim XY
	XY = findPicLocation("common\�Զ�.bmp", "�Զ�", False)
	If IsNull(XY) = False Then 
		MoveTo XY(0)+10, XY(1)+kOffsetY+10
		LeftClick 1
	End If
	
	Delay 1000
	Call SwitchGameWindow(1)
End Sub

// �����ʱ���ָ��ѡ�������μ�
Sub ClickTimeActivitiesOption(picPaths, name)
    Dim XY, i
    // �����ʱ�
    XY = findPicLocation("common\��ʱ�.bmp", "��ʱ�", True)
    Call HYLeftClick(XY(0), XY(1)+20)
    
    // ����ƶ�ѡ��
    Delay 1500
    Dim count 
    count = 3
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

/*
--------------------------------------------- ��� -----------------------------------------------
*/

Sub ��ӱ�(isEnd)
    // 1. ��ʱ��вμӡ���Ӹ�����
    Call Lib.HYOL.ClickTimeActivitiesOption("��Ӹ���\��Ӹ���1.bmp" & "|" & "��Ӹ���\��Ӹ���2.bmp", "��Ӹ���")
    Delay 2000
	
    Dim ��Ӹ�������X, ��Ӹ�������Y, XY, noTime
    ��Ӹ�������X = 0
    ��Ӹ�������Y = 0 

    For i = 0 To 13
        // �ж�ʣ������Ƿ�Ϊ0�� Ϊ0ʱ�˳��ű�
        noTime = Lib.HYOL.findPicLocation("��Ӹ���\��Ӹ���_ʣ�����Ϊ0.bmp", "��Ӹ���>ʣ�����Ϊ0", False)
        If IsNull(noTime) = False Then 
            If isEnd Then 
                Call Lib.HYOL.showEndMessage("ʣ�����Ϊ0����Ӹ��������", False)
            Else 
                Exit Sub
            End If
            
        End If
    
        // ������ѱ�, 40��δ�������쳣�˳�
        If ��Ӹ�������X = 0 Then 
            XY = Lib.HYOL.findPicLocationTimes("��Ӹ���\��Ӹ���_����.bmp", "��Ӹ���>����", 10, 2000, True)
            ��Ӹ�������X = XY(0)
            ��Ӹ�������Y = XY(1) - 100
        End If
        Call Lib.HYOL.HYLeftClick(��Ӹ�������X, ��Ӹ�������Y)
    
        // �����������ʾ���͡�ȷ�Ͻ��롱
        If i = 0 Then 
            Delay 1000
            Dim noTipLocation, confirmEnterLocation
            noTipLocation = Lib.HYOL.findPicLocation("��Ӹ���\��Ӹ���_���ε�¼������ʾ.bmp", "��Ӹ���>���ε�¼������ʾ", False)
            If IsNull(noTipLocation) = False Then 
                Call Lib.HYOL.HYLeftClick(noTipLocation(0) + 10, noTipLocation(1) + 10)
                Delay 500
            End If
            confirmEnterLocation = Lib.HYOL.findPicLocation("��Ӹ���\��Ӹ���_ȷ�Ͻ���.bmp", "��Ӹ���>ȷ�Ͻ���", False)
            If IsNull(confirmEnterLocation) = False Then 
                Call Lib.HYOL.HYLeftClick(confirmEnterLocation(0) + 30, confirmEnterLocation(1) + 20)
            End If
        End If
    
        // �������ս��
        Delay 6000
        Call Lib.HYOL.ClickFireButton
        Delay 3000

        // ������Զ����͡�2���١�
        If i = 0 Then 
            Call Lib.HYOL.ClickAutoFireButtonInAllWindow
            
        End If
    
        // Ԥ��ս��ʱ��
        Delay 30000
    
        // �����������Ϸ��
        XY = Lib.HYOL.findPicLocationTimes("��Ӹ���\������Ϸ.bmp", "��Ӹ���>������Ϸ", 30, 3000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 50, XY(1) + 20)
        Delay 4000
    Next
End Sub

/*
----------------------------------------- ������Ϸ --------------------------------------------
*/
Sub ������Ϸ(isEnd)
    Dim XY
    // ������� ������Ϸ
    XY = Lib.HYOL.findPicLocation("������Ϸ\������Ϸ.bmp", "������Ϸ", True)
    Call Lib.HYOL.HYLeftClick(XY(0)+20, XY(1)+20)
    Delay 3000

    // �ж��Ƿ������
    XY = Lib.HYOL.findPicLocation("������Ϸ\�����.bmp", "������Ϸ-�����", False)
    If IsNull(XY) = False Then 
        Goto ����Ƿ��д���
    End If

    Rem ��ʼ

    // �ж��Ƿ��Ѿ�ѡ���ˡ����ޡ�
    XY = Lib.HYOL.findPicLocation("������Ϸ\����_ѡ��.bmp", "������Ϸ-����_ѡ��", False)
    If IsNull(XY) Then 
        // �ж� ���� �Ƿ��Ѽ���
        XY = Lib.HYOL.findPicLocation("������Ϸ\����_����.bmp", "������Ϸ-����_����", False)
        If IsNull(XY) Then 
            XY = Lib.HYOL.findPicLocation("������Ϸ\����_����.bmp", "������Ϸ-����_����", True)
            Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
            Delay 1000
            XY = Lib.HYOL.findPicLocation("������Ϸ\��.bmp", "������Ϸ_��", True)
            Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
            Delay 2000
            XY = Lib.HYOL.findPicLocation("������Ϸ\����_����.bmp", "������Ϸ-����_����", True)
        End If
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
        Delay 1000
        XY = Lib.HYOL.findPicLocation("������Ϸ\��.bmp", "������Ϸ_��", True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
        Delay 2000
    End If

    Dim picName
    For i = 1 To 4
        If i = 1 Then 
            picName = "SS��"
        ElseIf i = 2 or i = 3 Then
            picName = "S��"
        Else 
            picName = "A��"
        End If
        // ��� SS
        XY = Lib.HYOL.findPicLocation("������Ϸ\" & picName & ".bmp", "������Ϸ_" & picName, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 20)
        Delay 10000

        // ��� ��ս
        XY = Lib.HYOL.findPicLocation("������Ϸ\��ս.bmp", "������Ϸ_��ս", True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 5000
	
        // ��� �Զ� �� 2����
        If i = 1 Then 
            Call Lib.HYOL.ClickAutoFireButton
        End If
        // ս��ʱ�����20��
        Delay 20000
    
        // ��� ȷ�� ��ť
        XY = Lib.HYOL.findPicLocationTimes("������Ϸ\ȷ��.bmp", "������Ϸ-ȷ��", 30, 3000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 2000
    Next
    Delay 3000
    // ���һ����ȡ
    XY = Lib.HYOL.findPicLocationTimes("������Ϸ\һ����ȡ.bmp", "������Ϸ-һ����ȡ", 3, 2000, True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 15)
    Delay 1000
    XY = Lib.HYOL.findPicLocation("������Ϸ\��.bmp", "������Ϸ-��", True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 20, XY(1) + 10)
    Delay 3000
    XY = Lib.HYOL.findPicLocation("common\ȷ��2.bmp", "������Ϸ-ȷ��2", True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
    Delay 1000

    Rem ����Ƿ��д���
    // �ж��Ƿ��д���
    XY = Lib.HYOL.findPicLocation("������Ϸ\����.bmp", "������Ϸ-����", False)
    If IsNull(XY) = False Then 
        Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 20)
        Delay 1000
        Goto ��ʼ
    End If
	
    If isEnd Then 
        Call Lib.HYOL.showEndMessage("������Ϸ���", False)
    End If

End Sub

/*
------------------------------------------------ �Ծ� ----------------------------------------------------
*/

Sub �Ծ�(isEnd)
    // 1. ��ʱ��вμӡ��Ծ���
    Call Lib.HYOL.ClickTimeActivitiesOption("�Ծ�\�Ծ�1.bmp" & "|" & "�Ծ�\�Ծ�2.bmp", "�Ծ�")

    Delay 5000
	
    Dim XY
    // 2. ������μ��Ծ���
    XY = Lib.HYOL.findPicLocationTimes("�Ծ�\�μ��Ծ�.bmp", "�μ��Ծ�", 5, 1000, True)
    Call Lib.HYOL.HYLeftClick(XY(0) + 40, XY(1) + 10)
    Delay 4000
	
    Dim matchBattleBtnX, matchBattleBtnY, ȷ�ϳ���X, ȷ�ϳ���Y, ȷ��X, ȷ��Y
    matchBattleBtnX = 0
    matchBattleBtnY = 0
    ȷ�ϳ���X = 0 
    ȷ�ϳ���Y = 0
    ȷ��X = 0
    ȷ��Y = 0


    // ʤ������С��12��, ��ͷ��С��96
    Dim i, ƥ�䵽ս��
    i=0
    While i<12
        // 3. �����ƥ��ս����
        If matchBattleBtnX = 0 Then 
            XY = Lib.HYOL.findPicLocationTimes("�Ծ�\ƥ��ս��.bmp", "�Ծ�-ƥ��ս��",5,2000,True)
            matchBattleBtnX = XY(0)
            matchBattleBtnY = XY(1)
        End If
        Call Lib.HYOL.HYLeftClick(matchBattleBtnX, matchBattleBtnY)
        Delay 2000
    
        // �ж��Ƿ�ƥ�䵽ս��
        Rem ��ѯ��ʼ
        ƥ�䵽ս�� = False
        For 10
            // �ж��Ƿ�ƥ��ʧ��
            XY = Lib.HYOL.findPicLocation("�Ծ�\ƥ��ս��.bmp", "�Ծ�-ƥ��ս��",False)
            If IsNull(XY) = False Then 
                Call Lib.HYOL.HYLeftClick(matchBattleBtnX, matchBattleBtnY)
                Delay 2000
                Goto ��ѯ��ʼ
            End If
            Delay 100
            // �ж϶Է��Ƿ��ѳ���
            XY = Lib.HYOL.findPicLocation("common\ȷ��.bmp", "�Ծ�-ȷ��", False)
            If IsNull(XY) = False Then 
                ƥ�䵽ս�� = True
                Exit For
            End If
		
            // �ж��Ƿ������ս��
            XY = Lib.HYOL.findPicLocation("common\�ٶ�_��ɫ.bmp", "�Ծ�>�ٶȻ�ɫ", False)
            If IsNull(XY) Then 
                Delay 10000
            Else 
                ƥ�䵽ս�� = True
                Exit For
            End If
        Next
        If ƥ�䵽ս�� = False Then 
            Call Lib.HYOL.showEndMessage("2������δƥ�䵽ս�����ű��˳�", True)
        End If
        Delay 1000
    
        // ����Զ�
        If i = 0 or i = 1 Then 
            Call Lib.HYOL.ClickAutoFireButtonInAllWindow
            Delay 1000
        End If
    
        // �ж��Ƿ��г��˰�ť���еĻ���ֱ�ӳ���
        XY = Lib.HYOL.findPicLocation("common\����.bmp", "�Ծ�-����", False)
        If IsNull(XY) Then 
            i = i + 1
            // ��ʱ10�룬�ж϶Է��Ƿ��г���
            Delay 10000
            XY = Lib.HYOL.findPicLocation("common\ȷ��.bmp", "�Ծ�-ȷ��", False)
            If IsNull(XY) Then 
                Delay 50000
            End If
        Else 
            // �������
            Call Lib.HYOL.HYLeftClick(XY(0), XY(1))
            Delay 1000
            If ȷ�ϳ���X = 0 Then 
                XY = Lib.HYOL.findPicLocation("common\ȷ�ϳ���.bmp", "�Ծ�-ȷ�ϳ���", True)
                ȷ�ϳ���X = XY(0)+20
                ȷ�ϳ���Y = XY(1)+10
            End If
            Call Lib.HYOL.HYLeftClick(ȷ�ϳ���X, ȷ�ϳ���Y)
            Delay 3000
        End If
    
        // �����ȷ�ϡ���ť
        XY = Lib.HYOL.findPicLocationTimes("common\ȷ��.bmp", "�Ծ�-ȷ��", 100, 5000, True)
        Call Lib.HYOL.HYLeftClick(XY(0) + 10, XY(1) + 5)
        Delay 5000
    
        // ���"ȷ��"��ť
        If ȷ��X = 0 Then 
            XY = Lib.HYOL.findPicLocation("common\ȷ��2.bmp|common\ȷ��.bmp", "�Ծ�-ȷ��", True)
            ȷ��X = XY(0) + 10
            ȷ��Y = XY(1) + 5
        End If
        Call Lib.HYOL.HYLeftClick(ȷ��X, ȷ��Y)
        Delay 2000
    Wend
    If isEnd Then 
        Call Lib.HYOL.showEndMessage("�Ծ�����", False)
    End If
End Sub

