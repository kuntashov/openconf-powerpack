$NAME ������ ��������� ������ �������������
'========================================================================================
'	1) ������ 1�-����������� �� ������������� � ����������� � ����������� �������,
'		������ ����� ���������� ��������� ����, 
' 		� ��� ������������� ������ � ����� ������ ���������� ������ �����
'	������. 
'	- ���� ������������ ��������� 1� � ����������� ������, � ���������� ��������������,
'		������ ���������� ����������� � ����������� ������.
'	- ���� ������������ ��������� 1� � ����������� ������, � � ���� ��� ���-�� ��������,
'		������ ���������� ����������� � ����������� ������.
'      
'		�������: RunInExclusiveMode � RunInSharedMode
'
'	���������� ��� ������� �������� �� F11 �  F10
'
'	2) �������� ������ ��� ������� ������������� �� ��������� ������  
'                                                           
'	����������. ��� ���������� ������ ������� ��������� ������������� �������� "1S.StatusIB.wsc"
'
'	Copyright (c) 2004 ����� �������")
'	e-mail: artbear@bashnet.ru 
'	ICQ: 265666057
'
'	������: $Revision: 1.14 $ 
'
'========================================================================================

' ��������� ��� ������ ��� ������ ���������� ���������
' ��� ��� ������ ����� � �������
Dim DebugFlag '����������� ���������� ����������
'DebugFlag = True '�������� ����� ���������� ���������

	sParamName = "/Open:"
  'RunFileName = BinDir+"Config\CmdLine\RunParseCmdLine.vbs"
  
Sub RunEnterpise(bInShared)

  Dim wshShell
  Dim fso 'as FileSystemObject

  set wshShell = CommonScripts.WSH
 
	CmdLine = CommandLine + " "
    Debug "1C CommandLine", CommandLine
	
	CmdLine = RegReplaceText(CmdLine, "\s+config\s+", " enterprise ")
 	if InStr(UCase(CmdLine), " ENTERPRISE ") = 0 then
		CmdLine = CmdLine + " enterprise "
	end if
	
	strExclusive = "/m"
	CmdLine = RegReplaceText(CmdLine, "\s+"+strExclusive+"\s+", " ")
	
	' ������� ���� � ����
	IBDirPath = CommonScripts.IBDir
	
 	if InStr(UCase(CmdLine), " /D") = 0 then
		CmdLine = CmdLine + " /D"+""""+IBDirPath+""""
	end if
           
	' ������� ��� ������������, ���� ��������� ������ ��������
	on Error resume next
	if InStr(UCase(CmdLine), " /N") = 0 then
		sUserName = AppProps(appUserName)
		CmdLine = CmdLine + " /N" + sUserName
	end if
	on Error goto 0

    Debug "CmdLine", CmdLine
	
	' ������ ����� ��������� ��������� ����
	bNeedRun = true
	bNeedSharedMode = bInShared
	
	set ob = CreateObject("1S.IBState")
	iIBState = ob.IBState(IBDirPath)

		'-1	-->>	"�������� ������������ ����� �� 1�, ����� ��������������"
		'0	-->>	"� ���� ������ ���"
		'1	-->>	"���� �������� ��� � ����������� ������ ��� ���� ������������ �������� � ����������� ������"
		'2	-->>	"1� �������� � ����������� ������"

	if (iIBState = -1) and bInShared then '
		' ������ ������ � ������������� �������������� � ����������� ������
		' TODO ? ��������� �������������� � �������� ������, � ����� �������� � ����������� ?
		strAnswer = "��������: " & ob.IBStateToString(iIBState) & vbCrLf & vbCrLf & _
				"�� ������ ����� � ����������� ����� � ��������������� ?"
		if MsgBox(strAnswer, vbOKCancel, SelfScript.Name) = vbOK then
			bNeedSharedMode = false
		else
			bNeedRun = false
		end if
	end if

	if (iIBState = 1) and not bInShared then
		' ������ ������ � ������������� ������ � ����������� ������ � �������� � ����������� ������
		strAnswer = "��������: " & ob.IBStateToString(iIBState) & vbCrLf & _
				"������ � ����������� ������ ����������" & vbCrLf & vbCrLf & _
				"�� ������ ����� � ����������� ������ ?"
		if MsgBox(strAnswer, vbOKCancel, SelfScript.Name) = vbOK then
			bNeedSharedMode = true
		else
			bNeedRun = false
		end if
	end if

	if (iIBState = 2) and not bInShared then
		' ������ ������ � ������������� ������ � ����������� ������ � �������� � ����������� ������
		strAnswer = "��������: " & ob.IBStateToString(iIBState) & vbCrLf & _
				"������ � ����������� ������ ����������" & vbCrLf & vbCrLf & _
				"�� ������ ����� � ����������� ������ ?"
		if MsgBox(strAnswer, vbOKCancel, SelfScript.Name) = vbOK then
			bNeedSharedMode = true
		else
			bNeedRun = false
		end if
	end if     
	
  	if not bNeedSharedMode then
		CmdLine = RegReplaceText(CmdLine, "\s+enterprise\s+", " enterprise " & strExclusive & " ")
	end if
    Debug "CmdLine", CmdLine
	
	if bNeedRun then
  		wshShell.Run CmdLine, 3, false
  	end if

End Sub ' RunEnterpise

' ��������� 1� �� � ����������� ������ ��� �������� ������������
' ���������� �������� �� F10
Sub RunInSharedMode()
	RunEnterpise true
End Sub ' RunInSharedMode

' ��������� 1� � ����������� ������ ��� �������� ������������
' ���������� �������� �� F11
Sub RunInExclusiveMode()
	RunEnterpise false
End Sub ' RunInExclusiveMode

' ������� ������� ����, ��������� � ��������� ������ ������������� ����� "/Open: "
Sub OpenExternalFileFromCommandLine()

	CmdLine = CommandLine
	Debug "CmdLine", CmdLine

	set CmdLineDict = CommonScripts.CommandLineToDictionary(CmdLine)
	arguments = CmdLineDict.Items
	For i = 0 To CmdLineDict.Count -1
		s = arguments(i)
		if InStr(LCase(s), LCase(sParamName)) = 1 then
			ErtName = Mid(s, Len(sParamName)+1)
			if CommonScripts.fso.FileExists(ErtName) then
				Debug "ErtName", ErtName

				Set ert = Documents.Open(ErtName)
				If ert Is Nothing Then      ' ������� ��������
					'If ert = docWorkBook Then   ' ��������, ��� ������� ������ �����
					CommonScripts.Error("�� ������� ������� ���� "+ErtName)
				End If
			end if
		end if
	Next

End Sub ' OpenExternalFileFromCommandLine

Sub Configurator_AllPluginsInit()
  OpenExternalFileFromCommandLine()
End Sub

Dim regEx ' Create variables.
Set regEx = New RegExp            ' Create regular expression.

Function RegReplaceText(srcStr, patrn, replStr)
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  RegReplaceText = regEx.Replace(srcStr, replStr)   ' Make replacement.
End Function ' RegReplaceText

Sub Debug(ByVal title, ByVal msg)
on error resume next
  DebugFlag = DebugFlag
  if err.Number<>0 then
    err.Clear()
    on error goto 0
    Exit Sub
  end if
on error goto 0

	CommonScripts.SetQuietMode(not DebugFlag)
	CommonScripts.Debug title, msg
End Sub'Debug

Sub EnableDebugMessages()
	DebugFlag = true
End Sub

Sub DisableDebugMessages()
	DebugFlag = false
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
