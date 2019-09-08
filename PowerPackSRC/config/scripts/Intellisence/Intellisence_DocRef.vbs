$NAME ����� ���������� ���/���
'/* ������������� ����� �� ������ ���������� ����������� ����������� ���� ��������� */
Dim CommonScripts
Dim doc			' ������������� ��������

Dim CountOfDots	' ���-�� �����, ��������� � �����. 
Dim StartCol	' ��������� �������
Dim BeginOfField '������� ������ ���� ����
Dim EndOfField '������� ����� ���� ����
Dim FieldBeforeIns ' �������� ��������, ��� �������� ���� ����
Dim TimerIns		' ������
 
Sub OpenList(param)
'	SendCommand 22500	' by ctrl-spc
	SendCommand 22510	' ������� ������ �������
End Sub

Sub Configurator_OnTimer(timerID)
    If timerID = TimerIns Then
	    Line1 = doc.SelStartLine
	    FieldAfterIns = doc.Range(Line1, BeginOfField, Line1, EndOfField)
	    If FieldBeforeIns <> FieldAfterIns Then
		    doc.Range(Line1, StartCol, Line1, BeginOfField) = ""
		    EndCol = StartCol+Len(FieldAfterIns)	' ������� ����� ����, ���� ��������� ������
		    doc.MoveCaret Line1, EndCol		' �������� ����
			CfgTimer.KillTimer TimerIns		' � ������� ������
	    End If
    End If
End Sub

Sub Telepat_OnInsert(InsertType, InsertName, Text)
	If CountOfDots <> 0 Then
		If CountOfDots < 2 Then
			Text = Text + "."
			BeginOfField = BeginOfField + Len(Text)
			CountOfDots = 2
			OpenList(0) 
		Elseif CountOfDots = 2 Then
		    Line1 = doc.SelStartLine
			EndOfField = BeginOfField + Len(Text)
    		FieldBeforeIns = doc.Range(Line1, BeginOfField, Line1, EndOfField)
    		CountOfDots = 0
			'������ �������� �� �������
			TimerIns = CfgTimer.SetTimer(10, False)
		End If
	End If
End Sub

Sub Attrs(MetaType)
    Set doc = CommonScripts.GetTextDoc(0)
    If doc Is Nothing Then Exit Sub

    Line1 = doc.SelStartLine
    Col1 = doc.SelStartCol
    Col2 = doc.LineLen(doc.SelStartLine)

    quStr=""""+MetaType+"."	' ������ �������
	BeginOfField = doc.SelStartCol + Len(quStr)

    CurrentText = doc.Range(Line1, Col1, Line1, Col2)
	CurrentText=+quStr+CurrentText

    doc.Range(Line1, Col1, Line1, Col2) = CurrentText

    doc.MoveCaret Line1, Col1+Len(quStr)
	CountOfDots = 1
	StartCol = Col1

	OpenList(0)
End Sub ' CodeReplace

Sub SprField()
	Attrs("����������")
End Sub

Sub DocField()
	Attrs("��������")
End Sub

' ������������� �������. param - ������ ��������,
' ����� �� ������� � �������
'
Sub Init(param)
    'Set CommonScripts = Scripts("Common")
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
	
    Set t = Plugins("�������")  ' �������� ������
    If Not t Is Nothing Then    ' ���� "�������" ��������
        ' ����������� ������ � �������� �������
        SelfScript.AddNamedItem "Telepat", t, False
        Telepat.DisableTemplateInRemString = 1      ' ��������� �������. 1-� ������������, 2-� �������
    Else
        ' ������ �� ��������. �������� � ������
        Scripts.Unload SelfScript.Name
    End If

	c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False

End Sub

' ��� �������� ������� �������������� ���
Init 0