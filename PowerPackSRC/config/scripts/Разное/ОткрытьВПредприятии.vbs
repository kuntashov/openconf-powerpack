'   ������ ��� ���������� �������� � ������ �����������.
'   
'�������� � ����� ������� (� �������������) ��������� ����� Open(), � ������� 
'������ ����� ������ � �����������. ���� ����������� ���� �������� �� �������������
'(�� F11 ��� ��������), �� ������ ����� ������ � �� ��� ��������.
'
'��������� �����:
'  - ������� ������
'  - ��������� �����
'  - ������ � ��������� ������������
'  - ����������� (����� ������� ����� ������)
'  - ������� ����������
'  
'��� ������ ��������� ��������� OpenIn1C.exe. Ÿ ����� �������� � ������� Bin, 
'���� � Bin\Config\System, ���� � �������, ����������� � PATH
'
' �����: ������� ����� aka ADirks
' e-mail: <adirks@ngs.ru>


Sub Open()
	set ob = CreateObject("1S.IBState")
	iIBState = ob.IBState(IBDir)
	If iIBState = -1 Then
		Set RunScript = Scripts("������ ��������� ������ �������������")
		Scripts("��������").Run1CInExlusiveMode()
	ElseIf iIBState = 0 Then
		Set RunScript = Scripts("������ ��������� ������ �������������")
		RunScript.RunInSharedMode()
	End If

	Set Doc = CommonScripts.GetTextDoc(true, 0)
	If Doc Is Nothing Then Exit Sub
	'Message Doc.Name
	'Message Doc.Path
	
	fname = BinDir & "OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\System\OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\Data\OpenIn1C\OpenIn1C.exe" ' ����, �.� � ���� ��� �������� Config\Sys - ����� ��� ����� � ����� �����
	If not CommonScripts.FSO.FileExists(fname) Then fname = "OpenIn1C.exe"


	
	PauseTime = 2000 'in milliseconds
	TypeLetter = ""
	if (Doc.ID >= 0) then ' �� ������� ����
		Set MetaObj = GetMetaObject(Doc.Name, TypeLetter)
		If Not MetaObj Is Nothing And TypeLetter <> "" Then
			parts = split(Doc.Name, ".")
			Repr = MetaObj.Name
			If MetaObj.Present <> "" Then Repr = MetaObj.Present
			If MetaObj.Descr <> "" Then Repr = Repr & " (" & MetaObj.Descr & ")"
			
			If parts(0) = "��������" Then Repr = "(������)" '��� ���������� ����� ������ ��������� ������ ������
			
			CommonScripts.RunCommand fname, """" & Repr & """ " & TypeLetter & " " & PauseTime, false
		End If
	Else
		CommonScripts.RunCommand fname, """" & Doc.Path & """ -f " & PauseTime, false
	End If
End Sub

Function GetMetaObject(Name, TypeLetter)
	Set GetMetaObject = Nothing
	
	parts = split(Name, ".")
	If UBound(parts) < 1 Then Exit Function
	
	On Error Resume Next
	Set GetMetaObject = MetaData.TaskDef.Childs("" & parts(0))("" & parts(1))
	On Error Goto 0
	If parts(0) = "����������" Then TypeLetter = "-s"
	If parts(0) = "������" Then TypeLetter = "-j"
	If parts(0) = "��������" Then TypeLetter = "-j"
	If parts(0) = "�����" Then TypeLetter = "-r"
	If parts(0) = "���������" Then TypeLetter = "-p"
End Function


'======= ������������� =============================
Private Sub Init()
	Set c = Nothing
	On Error Resume Next
	Set c = CreateObject("OpenConf.CommonServices")
	On Error GoTo 0
	If c Is Nothing Then
		Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
		Message "������ " & SelfScript.Name & " �� ��������", mInformation
		Scripts.UnLoad SelfScript.Name
	Else
		c.SetConfig Configurator
		SelfScript.AddNamedItem "CommonScripts", c, False
	End If
End Sub

Init
'===================================================
