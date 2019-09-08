'�����������.....
'�������� ��� �������� ����������� �� ������ ������.
'��������� ������ �� ������� ������ ������, �������� � ������ ��������� 
'(����, ���������, ���������, ��������������, IF) �� ������ ����� ��� ����. 
'��������� ������� � ������ �� ������� ������� (�� �����������) ������������. ����� ������,
'��� ������ ��������� ������������� ������ ������ ����� � ������.
'
'���� 3 �������� ���������;
'   1) ������� � �����������/���������� ��������� ����� - ������ GoUp, GoDown
'   2) ������� � ����������� ������ ����������� ��� � ���������� ��������� ����������� - ������ GoUp2, GoDown2
'   3) ������� � ������/���������� ������������ �����������. ��� ����, ���� ��� ������ ������������ 
'      ����������� ���������� ������/����� ���������, �� ������� ����.  ������ GoUp3, GoDown3
'
'������:
'<s1>��������� ������()
'	<s2>���� ... �����
'		.....
'		<s3>���� ... �����
'		<s4>���������;
'		<s5>���� ... ����
'			......
'			...<p1>...
'		<s6>����������;
'		.....
'		<p2>
'	���������
'��������������
'
'1) ��� �������� ����� �� ����� <p1> ��������������� �������� � ����� <s6>, <s5>, <s4>, <s3>, <s2>, <s1>
'2) ��� �������� ����� �� ����� <p1> ��������������� �������� � ����� <s5>, <s3>, <s2>, <s1>
'3) ��� �������� ����� �� ����� <p1> ��������������� �������� � ����� <s2>, <s1>.
'   ��� �������� �� ����� <p1> ��������������� �������� � ����� <s5>, <s2>, <s1>
'
'==================================================================================================
' ���� ����������� ���������� ��������� ���������:
'    ������ �. ��� trdm, 2004
'    ����� ������� aka artbear
'    ������� ����� aka ADirks
'==================================================================================================


Dim CommonScripts
Dim ArrStatement, ArrStatementUp, ArrStatementDn
Dim OpenParents, CloseParents
Dim ProcStarts, ProcEnds

Sub Init(dummy)
	Set CommonScripts = Nothing
	On Error Resume Next
	Set CommonScripts = CreateObject("OpenConf.CommonServices")
	On Error GoTo 0
	If CommonScripts Is Nothing Then
		Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
		Message "������ " & SelfScript.Name & " �� ��������", mInformation
		Scripts.UnLoad SelfScript.Name
	Else
		CommonScripts.SetConfig(Configurator)
	End If
	
	ArrStatement = Array("���������", "Procedure", "��������������", "EndProcedure", _
		"�������", 	"Function", "������������", "EndFunction", _
		"����", "If", "�����", "Then", "���������", "ElsIf", "�����", "Else", "���������", "EndIf", _
		"����", "While", "����", "Do", "����������", "EndDo", "���", "For", _
		"�������", "Try", "����������", "Except", "������������", "EndTry")

	ArrStatementUp = Array("���������", "Procedure", _
		"�������", 	"Function", "������������", "EndFunction", _
		"����", "If", "�����", "Then", "���������", "ElsIf", _
		"����", "While", "����", "Do", "���", "For", _
		"�������", "Try")
	ArrStatementDn = Array("��������������", "EndProcedure", _
		"�������", 	"Function", "������������", "EndFunction", _
		"���������", "ElsIf", "�����", "Else", "���������", "EndIf", _
		"����������", "EndDo", _
		"������������", "EndTry")

	OpenParents = Array("���������", "Procedure", _
		"�������", 	"Function", _
		"����", "If", _
		"����", "While", "����", "Do", "���", "For", _
		"�������", "Try")
	CloseParents = Array("��������������", "EndProcedure", _
		"������������", "EndFunction", _
		"���������", "EndIf", _
		"����������", "EndDo", _
		"������������", "EndTry")
		
	ProcStarts = Array("���������", "Procedure", "�������", "Function")
	ProcEnds = Array("��������������", "EndProcedure", "������������", "EndFunction")
End Sub

Init(0)
'=======================================================================


Sub GoUp()
	GoToStatement ArrStatement, -1, false
End Sub

Sub GoDown()
	GoToStatement ArrStatement, 1, false
End Sub

Sub GoUp2()
	GoToStatement ArrStatementUp, -1, false
End Sub

Sub GoDown2()
	GoToStatement ArrStatementDn, 1, false
End Sub

Sub GoUp3()
	GoToStatement ArrStatement, -1, true
End Sub

Sub GoDown3()
	GoToStatement ArrStatement, 1, true
End Sub

Sub Jump(nLine, nCol)
	TelepatJump = true
	Err.Clear
	On Error Resume Next
	Set Telepat = Plugins("�������")  ' �������� ������
	Telepat.Jump nLine, nCol, nLine, nCol, ModuleName
	If Err.Number <> 0 Then TelepatJump = false
	On Error Goto 0
	
	If Not TelepatJump Then
		Set Doc = CommonScripts.GetTextDocIfOpened()
		Doc.MoveCaret nLine, nCol
	End If
End Sub

Sub GoToStatement(ArrStatement, nDirection, bToOuterStatement)
    Set Doc = CommonScripts.GetTextDocIfOpened()
	If Doc Is Nothing Then Exit Sub
	If Doc.LineCount  = 0 Then Exit Sub

	nDirection = nDirection / Abs(nDirection)
	StartLine = Doc.SelStartLine + nDirection
	If nDirection < 0 Then
		EndLine = 1
	Else
		EndLine = Doc.LineCount
	End If
	Level = 0
	
	For nLine = StartLine To EndLine Step nDirection
		If bToOuterStatement Then
			Delta = AdjustLevel(Doc, nLine)
			Level = Level + Delta
			If Delta <> 0 Then
				FirstWord = GetFirstWord(Doc, nLine)
				'�� �������� ��������/������� ����� ��������������� ���������� �� ������
				If WordInArray(ProcStarts, FirstWord) OR WordInArray(ProcEnds,   FirstWord) Then
					'Doc.MoveCaret nLine, 0
					Jump nLine, 0
					Exit Sub
				End If
			End If

			If LevelHasChanged(nDirection, Level) Then
				Pos = CheckPlaceStatement(Doc, ArrStatement, nLine)
				'Doc.MoveCaret nLine, Pos-1
				Jump nLine, Pos-1
				Exit Sub
			End If
			
		Else
			Pos = CheckPlaceStatement(Doc, ArrStatement, nLine)
			IF Pos >= 0 Then
				'Doc.MoveCaret nLine, Pos-1
				Jump nLine, Pos-1
				Exit Sub
			End If
		End If
	Next
End Sub

'����������� �������� �����, � ���������� ��������� ������ �����������
Function AdjustLevel(Doc, nLine)
	AdjustLevel = 0
	Pos = CheckPlaceStatement(Doc, CloseParents, nLine)
	IF Pos >= 0 Then
		AdjustLevel = - 1
	Else
		Pos = CheckPlaceStatement(Doc, OpenParents, nLine)
		If Pos >= 0 Then
		AdjustLevel =  1
		End If
	End If
End Function

'� ����������� �� ����������� �������� ����������, ��������� �� ������ ������� �����������
Function LevelHasChanged(nDirection, Level)
	LevelHasChanged = false
	If nDirection < 0 Then
		If Level > 0 Then
			LevelHasChanged = true
		End If
	Else
		If Level < 0 Then
			LevelHasChanged = true
		End If
	End If
End Function

'����������, �������� �� ������ ����� � ������ ��������, � ���������� 
'��������� ������� ����� �����
'���� ����� �� ��������, �� ���������� -1
Function CheckPlaceStatement(Doc, ArrStatement, nLine)
	CheckPlaceStatement = -1

	FirstWord = GetFirstWord(Doc, nLine)
	If FirstWord = "" Then Exit Function

	sLine = Doc.Range(nLine)
	sLine = UCase(sLine)
	
	If WordInArray(ArrStatement, FirstWord) Then
		CheckPlaceStatement = InStr(1, sLine, FirstWord)
		Exit Function
	End IF	
End Function

'�������� ������ ����� � ��������� ������
Function GetFirstWord(Doc, nLine)
	GetFirstWord = ""

	sLine = Doc.Range(nLine)
	sLine = UCase(sLine)
	sLine = Replace(sLine, vbTab," ")
	sLine = Replace(sLine, ";"," ")
	sLine = Trim(sLine)
	ArrWord = Split(sLine)
	
	If UBound(ArrWord) >= 0 Then GetFirstWord = Trim(ArrWord(0))
End Function

'���������, ���� �� ����� � �������
Function WordInArray(Arr, Word)
	WordInArray = false
	For i = 0 To UBound(Arr) Step 1
		If UCase(Arr(i)) = Word Then
			WordInArray = true
			Exit Function
		End IF	
	Next
End Function