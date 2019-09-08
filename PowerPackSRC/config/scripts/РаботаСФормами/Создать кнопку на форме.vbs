'Attribute VB_Name = "Module1"
'      ������� ������ �� ����� c ������������ ��������
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'
'Option explicit
' ��������� ��� ������ ��� ������ ���������� ���������
' ��� ��� ������ ����� � �������
'Dim DebugFlag '����������� ���������� ����������
'DebugFlag = True '�������� ����� ���������� ���������

Dim fs
Dim IniDict

Dim ExistingButtonCaption, NewCaption, NewFormula, NewID

Sub SetIniDict(NullParam)
        Set IniDict = CreateObject("Scripting.Dictionary")
        IniDict.Add 1, "���������"
        IniDict.Add 2, "���"
        IniDict.Add 4, "����������X"
        IniDict.Add 5, "����������Y"
        IniDict.Add 6, "�����"
        IniDict.Add 7, "������"
        IniDict.Add 10, "����� ������� ������"
        IniDict.Add 12, "�������"
        IniDict.Add 13, "�������������"
        IniDict.Add 42, "����"
End Sub 'SetIniDict

Function ParseLineToDict(strLine)
        Dim Dict 'as Dictionary
        Set Dict = CreateObject("Scripting.Dictionary")
        Copy = Mid(strLine, 1)
        Index = 1
        Pos = CInt(InStr(Copy, ","))
        Do
                Current = Left(Copy, Pos - 1)
                Current = Mid(Current, 2, Len(Current) - 2) ' ����� �������'
                Key = Index
                If IniDict.Exists(Index) Then
                        Key = IniDict.Item(Index)
                End If
                Dict.Add Key, Current
                Copy = Mid(Copy, Pos + 1)

                Pos = CInt(InStr(Copy, ","))

                Index = Index + 1
        Loop Until Pos = 0
        If Pos = 0 Then
                Key = Index
                If IniDict.Exists(Index) Then
                        Key = IniDict.Item(Index)
                End If
                Copy = Mid(Copy, 2, Len(Copy) - 2)
                Dict.Add Key, Copy
        End If
        Set ParseLineToDict = Dict
End Function 'ParseLineToDict

Function DictToLine(Dict)
    a = Dict.Items
    s = ""
    For i = 0 To Dict.Count - 2
        s = s & """" & a(i) & """" & ","
    Next
    s = s & """" & a(Dict.Count - 1) & """" ' ��� �������
    DictToLine = s
End Function 'DictToLine

Function GetDialogLineForControl(DialogStream, ControlCaption, PositionBeginLine)
    GetDialogLineForControl = ""
'Stop
    PositionBeginLine = InStr(DialogStream, "{""" & ControlCaption & """") & ""
    If PositionBeginLine = 0 Then
        PositionBeginLine = InStr(DialogStream, "{""&" & ControlCaption & """") & ""
    End If
    If PositionBeginLine = 0 Then
        Message "������ ������ � ��������", mRedErr
        Exit Function
    End If

'    DebugEcho "PositionBeginLine", PositionBeginLine
    NextPos = InStr(Mid(DialogStream, PositionBeginLine), "},")

    GetDialogLineForControl = Mid(DialogStream, PositionBeginLine, NextPos - 1)

End Function

Function ReplaceDataInDialogLine(DialogLine, ExistingButtonCaption, NewCaption, NewFormula, NewID)
    Set Dict = ParseLineToDict(DialogLine)

    Dict.Item("���������") = NewCaption '"�������������"
    Dict.Item("�������") = NewFormula '"������������������_���.������������������(��������)"
    Dict.Item("�������������") = NewID '"�������������������"
    Dict.Item("����������X") = CInt(Dict.Item("����������X")) + CInt(Dict.Item("�����")) + 6
    ' ����� ��������� (�� ��������� - 54 �� 12 ���� "������������")
    Dict.Item("�����") = CInt(58 / 12 * Len(Dict.Item("���������")))

    NewDialogLine = DictToLine(Dict)
    NewDialogLine = "{" & NewDialogLine & "},"
    DebugEcho "NewDialogLine", NewDialogLine

    ReplaceDataInDialogLine = NewDialogLine
End Function ' ReplaceDataInDialogLine

'�������� ���������
Sub UnloadDoc(doc)

    Stream = doc.Stream
    DebugEcho "Stream", Stream

    StrCloseButton = GetDialogLineForControl(Stream, ExistingButtonCaption, Pos)
    DebugEcho "StrCloseButton", StrCloseButton

    strNewButton = ReplaceDataInDialogLine(StrCloseButton, ExistingButtonCaption, NewCaption, NewFormula, NewID)

'    strNewButton = "{" & strNewButton & "},"
'    DebugEcho "strNewButton", strNewButton

    InsertPos = Pos - 1

    Stream = Left(Stream, InsertPos - 1) & Chr(10) & strNewButton & Chr(10) & Mid(Stream, InsertPos + 1)

    doc.Stream = Stream + vbCrLf

    DebugEcho "Stream", doc.Stream

End Sub 'UnloadDoc'

Sub CreateButton(ExistingButtonCaption1, NewCaption1, NewFormula1, NewID1)
                                SetIniDict 0

        DebugEcho "ExistingButtonCaption1", ExistingButtonCaption1
        DebugEcho "NewCaption1", NewCaption1
        DebugEcho "NewFormula1", NewFormula1
        DebugEcho "NewID1", NewID1

        ExistingButtonCaption = ExistingButtonCaption1
        NewCaption = NewCaption1
        NewFormula = NewFormula1
        NewID = NewID1

        Dim w 'As CfgWindow
        Set w = Windows.ActiveWnd
        If Not w Is Nothing Then

            Dim d
            Set d = w.Document
            Set fs = CreateObject("Scripting.FileSystemObject")
'            If (d.ID = -1) And (d.Path <> "") And fs.FileExists(d.Path) Then ' ������� �����
            If (d.ID = -1) Then ' ������� �����
                UnloadDoc d.Page(0) ' ������
                Exit Sub
            ElseIf d.ID < 1 Then
                DebugEcho "���� �� �����, �� ������"
                Exit Sub
            End If

            If d = docText Then     ' ������ ������
'                AddDocument d.Name
            Else
                If d = docDEdit Then
                    DebugEcho d.Name, ""
                End If
                If d = docWorkBook Then ' �����
                    UnloadDoc d.Page(0) ' ������
                End If
            End If
        End If

End Sub 'CreateButton

Sub CreateButtonReset()
        CreateButton "�������", "�������������", "������������("+chr(34)+chr(34)+"�����"+chr(34)+chr(34)+"+����(35),, �����������������()); �����.�������(0);", "�������������������"
End Sub ' CreateButtonReset()

Sub CreateButtonMainFunctions()
        Formula = "����������������.�������_������������(��������, " & """""""""" & ");"
        DebugEcho "Formula", Formula
        CreateButton "�������", "����������...", Formula, "��������������"
End Sub ' CreateButtonMainFunctions()

Sub CreateSomeButton()
        sButtonName = InputBox("������� �������� ������:")
        If sButtonName = "" Then
                Exit Sub
        End If
        sButtonFormula = InputBox("������� ������� ������:")
        If sButtonFormula = "" Then
                Exit Sub
        End If
        If InStr(sButtonFormula, "(") = 0 Then
            sButtonFormula = sButtonFormula + "()"
        End If
        sButtonID = InputBox("������� ������������� ������ (�������������):")

        CreateButton "�������", sButtonName, sButtonFormula, sButtonID
End Sub 'CreateSomeButton

Sub Echo(text)
        Message text, mNone
End Sub 'Echo

Sub Debug(title, msg)
On Error Resume Next
  DebugFlag = DebugFlag
  If Err.Number <> 0 Then
    err.Clear()
    On Error GoTo 0
    Exit Sub
  End If
  If DebugFlag Then
    If Not (IsEmpty(msg) Or IsNull(msg)) Then
      msg = CStr(msg)
    End If
    If Not (IsEmpty(title) Or IsNull(title)) Then
      title = CStr(title)
    End If
    If msg = "" Then
      Echo (title)
    Else
      Echo (title + " - " + msg)
    End If
  End If
On Error GoTo 0
End Sub 'Debug

Sub DebugEcho(title, msg)
        Debug title, msg
End Sub 'DebugEcho

Sub DebugVar(NameVar)
On Error Resume Next
  Dim a
  a = Eval(NameVar)
  If Err <> 0 Then
    err.Clear()
    Exit Sub
  End If
  If IsNull(a) Or IsEmpty(a) Then
    call Debug(NameVar,"null")
  Else
    call Debug(NameVar,a)
  End If
On Error GoTo 0
End Sub 'DebugVar
'
'Set Wrapper = CreateObject("DynamicWrapper")
'Wrapper.Register "USER32.DLL", "GetFocus", "f=s", "r=l"
'Wrapper.Register "USER32.DLL", "SendMessage", "I=lllr", "f=s", "r=l"
'Set CommonScripts = scripts("common")
Function FindWindow(MainWindowText, ChildWindowText, WaitingTime)
  Dim retval
  retval = False
  'statements here
  Dim objFindWindow
  Set objFindWindow = CreateObject("ArtWin.Win")
  Dim bFindWindow, hFindWindow
  Dim iSleepTime, iSleepCounter
  iSleepTime = 300

  Debug "������ ������ ����", MainWindowText+", "+ChildWindowText
  Debug "����� ��������", Round(WaitingTime/1000)
  bFindWindow = False
  FindWindow = False ' ver 2.4
' For iSleepCounter = 1 To Int(5000/iSleepTime) 'Step step
  'For iSleepCounter = 1 To Round(WaitingTime / iSleepTime) '6 �����
'    wScript.sleep iSleepTime
'   Debug "SleepCounter", iSleepCounter
    If (ChildWindowText = "") Then
      bFindWindow = CBool(objFindWindow.Find(MainWindowText))
    Else
      hFindWindow = objFindWindow.FindChild(MainWindowText, ChildWindowText)
      bFindWindow = False
      If hFindWindow <> 0 Then
        bFindWindow = True
      End If
    End If
    If bFindWindow Then
      'Exit For
    End If
  'Next 'iSleepCounter
  If Not bFindWindow Then
    Debug "���������� �� �������� - ����� ���� "+MainWindowText+", " _
      + ChildWindowText,""
    Set objFindWindow = Nothing
    Exit Function
  Else
'   Debug "����� ���� == ",MainWindowText+", "+ ChildWindowText
  End If

  Set objFindWindow = Nothing
  FindWindow = True
End Function 'FindWindow
Function GetStream()

End Function 'FindWindow

Function GetFocusWindowText()
    GetFocusWindowText = ""

	set wshShell = createObject("WScript.Shell")
    'wshShell.SendKeys("{ENTER}")
    'wshShell.SendKeys("{TAB}{TAB}{TAB}{TAB}")
    wshShell.SendKeys("^TAB")
    'wshShell.SendKeys("^TAB")

	NumberOfTrying = 1
	StartFormulaTimer(500)
End Function

'Dim NumberOfTrying
NumberOfTrying = 0
TimerFormula = 0
Sub StartFormulaTimer(Interval)
    'StopFormulaTimer
    'TimerFormula = CfgTimer.SetTimer(Interval, False)
End Sub
Sub StopFormulaTimer()
    'If TimerFormula <> 0 Then CfgTimer.KillTimer TimerFormula
End Sub

'Sub Configurator_OnTimer(timerID)
Sub Configurator_OnTimer1(timerID)
    If timerID = TimerFormula Then
		StopFormulaTimer()
		if NumberOfTrying = 1 then
			set wshShell = createObject("WScript.Shell")
		    'wshShell.SendKeys("^{TAB}{TAB}{TAB}{TAB}{TAB}")
		    'wshShell.SendKeys("^{TAB}")
		    'wshShell.SendKeys("^{TAB}{TAB}{TAB}{TAB}{TAB}")
			NumberOfTrying = 2
			StartFormulaTimer(500)
			Exit Sub
		end if
		'    ActiveWindowText = Windows.ActiveWnd.hWnd
		'    ActiveWindowText = Windows.ActiveWnd.Caption
		'Echo ActiveWindowText
		'    fFindWindow = FindWindow(ActiveWindowText, "�&�����������", 10000)
		'    fFindWindow = FindWindow(ActiveWindowText, "1", 10000)
		'Echo CStr(fFindWindow)
		'Exit Function
		'    'Handle = ActiveWindow.Caption

		  Dim Wrapper
		Set Wrapper = CreateObject("DynamicWrapper")
		Wrapper.Register "USER32.DLL", "GetFocus", "f=s", "r=l"
		Wrapper.Register "USER32.DLL", "SendMessage", "I=lllr", "f=s", "r=l"
		Wrapper.Register "USER32.DLL", "GetTopWindow", "I=l", "f=s", "r=l"
		Wrapper.Register "USER32.DLL",   "GetForegroundWindow",      "f=s", "r=l"
		'Set CommonScripts = scripts("common")

		  Dim objFindWindow
		  Set objFindWindow = CreateObject("ArtWin.Win")

		Echo "1"
		    'for i=1 to 1000000
			'next
		'Echo "2"

		PropWndCaption = Space(128)
		PropWnd = Wrapper.GetForegroundWindow '  Foreground ����
		Echo "3"
		tcnt = Wrapper.SendMessage (PropWnd, &HD ,128, PropWndCaption) 'WM_GETTEXT
		Echo "4"
		Echo "tcnt <" & CStr(PropWndCaption) & ">"
		'If InStr(UCase(CStr(PropWndCaption)),"��������") = 1 then '����� �������� �� ������� ������ Foreground ����
		'	Wrapper.SendMessage PropWnd, &H10 ,NULL, NULL 'WM_CLOSE
		'End If

		    Buff = Space(128)
		    ActiveControl = Wrapper.GetFocus '������� ���� ������� � ������� �����
		Echo "ActiveControl <" & CStr(ActiveControl) & ">"

		    text = CStr(objFindWindow.GetWindowText(ActiveControl))
		Echo "<" & text & ">"

		    ActiveControl = Wrapper.GetTopWindow(Windows.ActiveWnd.hWnd) '������� ���� ������� � ������� �����
		Echo "ActiveControl <" & CStr(ActiveControl) & ">"

		    text = CStr(objFindWindow.GetWindowText(ActiveControl))
		Echo "<" & text & ">"

		    ActiveControl = Wrapper.GetTopWindow(ActiveControl) '������� ���� ������� � ������� �����
		Echo "ActiveControl <" & CStr(ActiveControl) & ">"

		    text = CStr(objFindWindow.GetWindowText(ActiveControl))
		Echo "<" & text & ">"

		    ActiveControl = Wrapper.GetTopWindow(ActiveControl) '������� ���� ������� � ������� �����
		Echo "ActiveControl <" & CStr(ActiveControl) & ">"

		    text = CStr(objFindWindow.GetWindowText(ActiveControl))
		Echo "<" & text & ">"

		Set d = Windows.ActiveWnd.Document
		    If d = docWorkBook Then ' �����
		        Set book = d.Page(0) ' ������
		    End If

		    cnt = Wrapper.SendMessage(ActiveControl, &HD, 128, Buff)  ' WM_GETTEXT
		    If cnt > 0 Then ' ���� � ������� ��� - ������ �������
		        Buff = Left(CStr(Buff), cnt)
		Echo Buff
		        GetFocusWindowText = Buff
		    End If
    End If
End Sub

Sub GoToFormula()
    Buff = Space(128)
    ActiveControl = Wrapper.GetFocus '������� ���� ������� � ������� �����
    cnt = Wrapper.SendMessage(ActiveControl, &HD, 128, Buff)  ' WM_GETTEXT
    If cnt > 1 Then ' ���� � ������� ��� - ������ �������
        Buff = Left(CStr(Buff), cnt)
        If InStr(Buff, "#") > 0 Or InStr(Buff, "?(") > 0 Or InStr(Buff, "(") = 0 Then
            Message "�������� ��� ���������/�������."
        Exit Sub
        End If
        Set doc = CommonScripts.GetTextDocIfOpened(1)

        ModuleText = Split(doc.text, vbCrLf)

        doc.MoveCaret 0, 0 '� ������ ������

        For i = 0 To UBound(ModuleText) '������ � ������ ���������/�������
            sText = UCase(ModuleText(i))
            ProcName = UCase(Left(Buff, Len(Buff) - 1)) '��� ���������/������� ��� ������ ������
            If InStr(sText, "��������� " & ProcName) = 1 Or InStr(sText, "������� " & ProcName) = 1 Then
                doc.MoveCaret i, 0
                Exit Sub ' ����� ���������/�������, �����
            End If
        Next

        ' ���� �� ����� ����� �������� �

        ProcOrFunc = InputBox("��� ���������: 1 - ���������, 2 - ������� ", "������� ��������", "1")

        If ProcOrFunc = "1" Then
            ReplValue1 = "��������� ": ReplValue2 = "�������������� // "
        ElseIf ProcOrFunc = "2" Then
            ReplValue1 = "������� ": ReplValue2 = "������������ // "
        Else
            Exit Sub
        End If

        doc.MoveCaret UBound(ModuleText), 0 '��������� ���� ���������
        For i = doc.SelStartLine To 0 Step -1 '����� ��������� ���������/�������
            sText = UCase(ModuleText(i))
            If InStr(sText, "��������������") = 1 Or InStr(sText, "������������") = 1 Then
                doc.MoveCaret i, 0
                Exit For
            End If
        Next
        doc.MoveCaret i + 1, 0

        iText = vbCrLf & "//----------------------------------------------------------" & vbCrLf & _
        ReplValue1 & Buff & vbCrLf & vbCrLf & _
        ReplValue2 & Buff
        doc.range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = iText
    End If
End Sub
