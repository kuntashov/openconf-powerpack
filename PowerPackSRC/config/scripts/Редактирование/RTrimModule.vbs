'$ENGINE {B54F3741-5B07-11CF-A4B0-00AA004A55E8}

'Author: Valera e-mail: wash_ai@mail333.com
'������ ��� �������� �������� � ������ ����� ����������� �������
'����� ���� �� ���������� ������ ������ ������ ����� �������� ������ ����
'��������� ��������� ��������� MD-���� � ������� ���������
'��������� ������������������ �������� ��������� ������

Dim NumOfStrings

'��������� ������ �������� � ������� ���������� ���������
Function TrimThisDoc(doc, mode)
	Dim TextStrings
	TextStrings = Split(doc.Text, Chr(13) & Chr(10))
	s1 = 0
	s2 = 0
	x = UBound(TextStrings, 1)
	For i = 0 To x
		NumOfStrings = NumOfStrings + 1
		T = TextStrings(i)
		s1 = s1 + Len(T)
		'If mode = "All" Then
			' �� ���� ����� ������ ���� ��� ��������� ������������� �������+���������
            ' �.�. ���� �����:
            '
            'prevLen = Len(T) + 2
            'Do While prevLen <> Len(T)
            ' prevLen = Len(T)
            ' T = Replace(T, Chr(32) & Chr(32), Chr(32))
            ' T = Replace(T, Chr(9) & Chr(9), Chr(9))
            ' T = Replace(T, Chr(32) & Chr(9), Chr(9))
            ' T = Replace(T, Chr(9) & Chr(32), Chr(9))
            'Loop
            '
            ' ��� ����������� �������, ��!!!! ���� � ��� ������������ ���-������ � ���� ���������� ����������
            ' ������ �� ���������� ��������, ��, ���, ������ ����� ������������� ������ ����
            ' ��� ��� ����� ����� ������ ���, ������� �� ���������� ��������� ������� � �������� ��, ��� �/� ����
            ' ���� ��� ��� ������ ���� - ����� �����, � ����� � ��� ������ ���������...
        'End If
        If (mode = "Right") Or (mode = "All") Then
            flag = 0
            Do While (Len(T) > 0) And (flag = 0)
                If flag = 0 Then
                	If (Right(T, 1) = Chr(9)) Or (Right(T, 1) = Chr(32)) Then
                        T = Left(T, Len(T) - 1)
                    Else
                        flag = 1
                    End If
                End If
            Loop
            T = RTrim(T)
        End If
        If (mode = "Left") Or (mode = "All") Then
            flag = 0
            Do While (Len(T) > 0) And (flag = 0)
                If flag = 0 Then
                    If (Left(T, 1) = Chr(9)) Or (Left(T, 1) = Chr(32)) Then
                        T = Right(T, Len(T) - 1)
                    Else
                        flag = 1
                    End If
                End If
            Loop
            T = LTrim(T)
        End If
        s2 = s2 + Len(T)
        TextStrings(i) = T
    Next
    doc.Text = Join(TextStrings, Chr(13) & Chr(10))
    prevLen = Len(doc.Text) + 2
    Do While prevLen <> Len(doc.Text)
        prevLen = Len(doc.Text)
        doc.Text = Replace(doc.Text, Chr(13) & Chr(10) & Chr(13) & Chr(10), Chr(13) & Chr(10))
    Loop
    TrimThisDoc = s1 - s2
End Function

'������������������ ��������� ������
Sub ReFormatCurrentWnd()
    NumOfStrings = 0
	
	Set d = CommonScripts.GetTextDocIfOpened(false) ' �� ������������ �������� ������
	If d is Nothing Then
		Exit Sub
	End If	
	
	docName = d.Name
	s0 = TrimThisDoc(d, "All")
	d.MoveCaret 0, 0, d.LineCount	
    d.FormatSel
    'Message "���������� �����: " & NumOfStrings, mInformation
End Sub

'�������� �������� � ������� �������� ������
Sub RTrimCurrentWnd()
    NumOfStrings = 0
	
	Set d = CommonScripts.GetTextDocIfOpened(false) ' �� ������������ �������� ������
	If d is Nothing Then
		Exit Sub
	End If	
    
	s0 = TrimThisDoc(d, "Right")
    docName = d.Name
	
    Message "���������� �����: " & NumOfStrings, mInformation
    If s0 = 0 Then
        Message "�� ������� ������ �������� � " & docName, mInformation
    Else
        Message "� ������ '" & docName & "' ������� ������ ������� ������ ���������� �� " & s0, mInformation
    End If
End Sub

'�������� �������� � ���������� ������, � ����������, � ����� �� ���������� ������� � ����������
Sub RTrimGlobalModuleDocReportsAndCalcVars()
    NumOfStrings = 0
    s0 = 0
    Set doc = Documents("����������������.������")
    s0 = s0 + TrimThisDoc(doc, "Right")
    Set refs = MetaData.TaskDef.Childs("��������")
    For i = 0 To refs.Count - 1
        Message refs(i).Name, mInformation
        Set doc = Documents("��������." & refs(i).Name & ".�����.������")
        s0 = s0 + TrimThisDoc(doc, "Right")
        Set doc = Documents("��������." & refs(i).Name & ".������ ���������.������")
        s0 = s0 + TrimThisDoc(doc, "Right")
    Next
    Set refs = MetaData.TaskDef.Childs("�����")
    For i = 0 To refs.Count - 1
        Set doc = Documents("�����." & refs(i).Name & ".�����.������")
        s0 = s0 + TrimThisDoc(doc, "Right")
    Next
    Set refs = MetaData.TaskDef.Childs("���������")
    For i = 0 To refs.Count - 1
        Set doc = Documents("���������." & refs(i).Name & ".�����.������")
        s0 = s0 + TrimThisDoc(doc, "Right")
    Next
    Message "���������� �����: " & NumOfStrings, mInformation
    Message "����� ������� " & s0 & " ��������", mInformation
End Sub

'
' ��������� ������������� �������
'
Sub Init(dummy) ' ��������� ��������, ����� ��������� �� �������� � �������
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
        Message "������ " & SelfScript.Name & " �� ��������", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub
 
Init 0 ' ��� �������� ������� ��������� �������������

