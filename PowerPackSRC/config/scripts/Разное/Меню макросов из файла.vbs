$NAME ���� �������� �� �����
'      ������ ��� ���������� �������� �� �������� � ���� ������ �� Ctrl-2 (c) artbear,  2004
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057

' ������ �������� � ���� � �������� "�������"
' � ������������ ������� �� ����
' ���������� � bin\config\scripts
'
' ��� ���������� ������ � �������� bin\config ������ ���� ���� Macros.ini
' c ������������ ������� �� �������
' ������ �����:
' ����������.���������� = ����� ������������ �������, ������� �������� � ����
' ����� ������������ ������ "-" �� ��������� ������,  ��� ����� ����������� ����
' ���� ������� [-],  ����� ����� ����������� ����� ��������
'
' ������ �������� � ���������� ������� [���������],  [������ � �������]
'
' �����������, ��� ������� ����� ��������� � ���������� �������
' ��������,
' 1C++.OpenDefClsPrm = ������� ���� ����������� 1�++ (defcls.prm)
'
' ���������� ����������� ���� Macros.ini - ��� ��� ����
'
Dim gsMacrosFileName ' ���������� �� ��������� � bin\config
  gsMacrosFileName = "config\Macros.ini"

  sConstMenuPrefix = "���� �������� �� �����"

Dim Telepat

'  Dim ResDict 'as Dictionary
Dim MenuForTelepat ' as string

' ���������� ������� "����� ���� ��������"
' ��������� ����������� �������� ������� � ���� ��������.
' ��� ����� ������ ������� ������-��������� ����������� �������.
' ������ ����������� ����� ������ ������������� �� ��������� ������.
' ��������� ������ ������ ������������ �� ��������� � ������ ������.
' ����� �������� ������� ����� ������ | ����� ������� �������
' ����� ����� ������ | ����� ������� ������������� �������.
' � ���� ������ � ������� OnCustomMenu ������ �������� ������ ����
' ����� ������� ���� �������������
' ��� �������� ������������ ������� ��� "-"
'
Function Telepat_GetMenu()   
  'Telepat_GetMenu = Telepat_GetMenu & MenuForTelepat
  RefreshMacros
  Telepat_GetMenu = MenuForTelepat
End Function ' Telepat_GetMenu

' ���������� ������� OnCustomMenu.
' ���������� ��� ������ ������������� ������ ����,
' ������������ � "GetMenu"
' Cmd - �������� (��� �������������) ���������� ������ ����.
'
Sub Telepat_OnCustomMenu(CmdArtur)
	Dim M
	s = Trim(CStr(CmdArtur))
	If (InStr(s, sConstMenuPrefix) = 1) Then
		On Error Resume Next
		' ���������� ������������� ������ ��� ������ �������� �� ��������
		' (����� � �������������� Eval �� �������� ������������ js: ��, ��� �
		' js ��� �������� � ����������� ������� ��������������� �� ������� ����� -- a13x
		MacrosSpec = Replace(CmdArtur, sConstMenuPrefix, "")
		M = Split(MacrosSpec, "::", 2)
		CommonScripts.CallMacros M(0), M(1)
		On Error GoTo 0
	End If
End Sub ' Telepat_OnCustomMenu

' �������� ������ �� INI-�����
Function GetDataFromIniFile(ByVal IniFileName, ByVal FlagInsertMenuGroup)
	GetDataFromIniFile = False

	' ����� �������������
	Dim IniFile 'As TextStream

	On Error Resume Next
	Dim ForRead
	ForRead = 1

	Dim fso 'as FileSystemObject
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set IniFile = fso.OpenTextFile(IniFileName, ForRead)
	If Err.Number <> 0 Then
		Err.Clear()
		CommonScripts.Echo "Ini-���� " & IniFileName & " �� ������� �������!"
		Exit Function
	End If
	On Error GoTo 0

	'Set ResDict = CreateObject("Scripting.Dictionary")

	Dim reg 'As RegExp
	Set reg = New RegExp
		reg.Pattern = "^\s*([^\.]+)\.([^=]+)\s*=\s*([^;']+)[;']?"
		reg.IgnoreCase = True

	Dim regForGroup 'As RegExp
	Set regForGroup = New RegExp
		regForGroup.Pattern = "^\s*\[([^\]]+)\]\s*[;']?"
		regForGroup.IgnoreCase = True

'    Dim regForSeparator 'As RegExp
'    Set regForSeparator = New RegExp
'        regForSeparator.Pattern = "^\s*\[-\]\s*[;']?"
'        regForSeparator.IgnoreCase = True

	If FlagInsertMenuGroup Then
		MenuForTelepat = "������� � ������� ���������" & vbCrLf
	Else
		MenuForTelepat = ""
	End If
	TabForGroup = ""

	bFind = False
	Do While IniFile.AtEndOfStream <> True
		s = Trim(IniFile.ReadLine)
	' ���� �� ������-�����������
		If Not RegExpTest("^\s*[;']", s) Then
		If Left(s, 1) = "-" Then ' �����������
'        If RegExpTest("^\s*\[-\]\s*", s) Then ' �����������
		  MenuForTelepat = MenuForTelepat & TabForGroup & vbTab & "-" & vbCrLf
		Else
		Set Matches = regForGroup.Execute(s) ' ��� ������ ����
		If Matches.Count > 0 Then
	'        CommonScripts.Echo "������ " & Matches(0).SubMatches(0)
		    If RegExpTest("^\s*\[-\]\s*", s) Then ' �����������
		        MenuForTelepat = MenuForTelepat & "-" & vbCrLf
		    Else
		        If FlagInsertMenuGroup Then
		            TabForGroup = vbTab
		        Else
		            TabForGroup = ""
		        End If
		        MenuForTelepat = MenuForTelepat & TabForGroup & Matches(0).SubMatches(0) & "| | " & vbCrLf
		    End If
		Else
	' �������� ���� � �������� � Ini-�����, ����� ��������� �����������
		  Set Matches = reg.Execute(s)
		  If Matches.Count > 0 Then
	' �������� ����� ����, �������� �� �������� ��������� � �����(� ������) �������
		    sScriptName = Trim(Replace(Matches(0).SubMatches(0), vbTab, " "))
		    sMacrosName = Trim(Replace(Matches(0).SubMatches(1), vbTab, " "))
		    sMacrosReplaceName = Trim(Replace(Matches(0).SubMatches(2), vbTab, " "))
			
			' a13x
			sEvalMacrosName = sScriptName + "::" + sMacrosName

		    If CommonScripts.MacrosExists(sScriptName, sMacrosName) Then
	'            ResDict.Add sMacrosReplaceName, sEvalMacrosName
		        bFind = True

		        MenuForTelepat = MenuForTelepat & TabForGroup & vbTab & sMacrosReplaceName & "| |"
				MenuForTelepat = MenuForTelepat & sConstMenuPrefix & sEvalMacrosName & vbCrLf
		    End If
		  End If
		End If
		End If
		End If
	Loop
	IniFile.Close()

	'if ResDict.Count=0 then ' �� ����� �� ������ �����
	If bFind = False Then ' �� ����� �� ������ �����
		CommonScripts.Echo "�� ������� �������� ������ �� Ini-����� " & IniFileName
		GetDataFromIniFile = False
	Else
		GetDataFromIniFile = True
	End If
End Function ' GetDataFromIniFile

' ��������� �� ������������ �������
' ������� �������� �� �����
Dim regExTest
Function RegExpTest(ByVal patrn, ByVal strng)
	If IsEmpty(regExTest) Then
		Set regExTest = New RegExp         ' Create regular expression.
	End If
	regExTest.Pattern = patrn         ' Set pattern.
	regExTest.IgnoreCase = True      ' disable case sensitivity.
	RegExpTest = regExTest.Test(strng)      ' Execute the search test.
End Function ' RegExpTest

Function RefreshMacros()
	RefreshMacros = True	
	GetDataFromIniFile BinDir & gsMacrosFileName, False 'True
End Function ' RefreshMacros

Function OpenParametersFile()
	OpenParametersFile = True	
	Documents.Open BinDir & gsMacrosFileName
End Function ' OpenParametersFile

' �������� ��� ������� � ���� "config\Macros_all.ini",
' ��� �� ����� ����������� � ������������� � ����
Function SaveAllMacrosToFile()
	
	SaveAllMacrosToFile = True

  	sUnloadFileName = "config\Macros_all.ini"

	Set fso = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	Set IniFile2 = fso.CreateTextFile(BinDir & sUnloadFileName, True)
	On Error GoTo 0
	
	Set e = CreateObject("Macrosenum.Enumerator")
	iScriptCount = Scripts.Count - 1
	For i = 0 To iScriptCount
		sScriptName = Scripts.Name(i)
		bInsertScript = 0
		Set script = Scripts(i)
		arr = e.EnumMacros(script) ' ��������� ������� �������� �������
		For j = 0 To UBound(arr)
			IniFile2.WriteLine sScriptName & "." & arr(j) & " ="
		Next
	Next

End Function ' SaveAllMacrosToFile

' ����� � ������ ������ ��� �������, ����� ������ ��,
' ��� ������ ����������� �� �������� ����� �������
Sub Configurator_AllPluginsInit()
	'Init 0
	RefreshMacros
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
	
	' ��� �������� ������� �������������� ���
  	c.AddPluginToScript SelfScript, "�������", "Telepat", Telepat
  
	SelfScript.AddNamedItem "CommonScripts", c, False

	RefreshMacros
End Sub

Init 0

Sub ShowStructuredMenuOfMacros()
	GetDataFromIniFile BinDir & gsMacrosFileName, False	
	Set srv = CreateObject("Svcsvc.Service")
	Telepat_OnCustomMenu srv.PopupMenu(MenuForTelepat, 0)
End Sub
