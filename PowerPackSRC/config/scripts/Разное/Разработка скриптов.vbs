$NAME ���������� ��������
' ��� ������������� ��������
'
' ������: $Revision: 1.14 $
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057

' ��� ������������� �������� ���� ����� ������ ������ "���������� ��������",
' ������� ��������� ���� ��� ������ ������������ ������,
' � � ���������� ����� ������ ������ ������������� ������ ������ � �������� ��� �������.
' ���� �������� ���������, ����� ����� ������� ������ ��� ������ �������,
' ���� ������ ����� ����, �� ����� �� � ����������.
' ��� ����: ���� �������/����������� ������ ��� ������ � VB,
' �� ��� ������ ��������� ������ ������� "Attribute VB_Name",
' ������� ��������� VB � �� ������� �������� ������������.
'
' ������ "SelectActionForScript" ������ ������ ���� ����������� ��������
'  	� ��� ���������� ������� ��������� ��������� ������������ �������� (������� �����, ������������� ������, ��������� ������)
'
Dim sRepositoryPath ' ���� � ����������� �������� ( ������ ����������� ������������� ��� ���� ���� ���� )
	sRepositoryPath = "T:\����������������\OpenConf\����������� ��������\OpenConf_Scripts\�������" 

Dim sNameDevelopScript
Dim sDevelopScriptPath

Dim iPrevSelectedScriptIndex ' 
	iPrevSelectedScriptIndex = -1

' ������������� ������, ������� ����� �������������
' ����� �� ��� �������������
Sub SetupScriptForDevelopmentInner(sNameDevelopScriptParam)

	sNameDevelopScript = sNameDevelopScriptParam

    If sNameDevelopScript = "" Then
        Exit Sub
    End If

    sDevelopScriptPath = sNameDevelopScript
    If (InStrRev(UCase(sNameDevelopScript), ".VBS") > 0) Or (InStrRev(UCase(sNameDevelopScript), ".JS") > 0) Then

		' TODO $NAME �� ����������� ����� � ������ ������, ����� ���������� ������ �� ������ 
		' �� ������ ������, ����, ������������ ����������� � $ENGINE
		FirstLine = CommonScripts.FSO.OpenTextFile(sDevelopScriptPath).ReadLine()
		If Left(FirstLine, 5) = "$NAME" Then
			sNameDevelopScript = Trim(Mid(FirstLine, 6))
		Else			
	        sNameDevelopScript = LCase(sNameDevelopScript)
    	    sNameDevelopScript = Replace(sNameDevelopScript, ".vbs", "")
        	sNameDevelopScript = Replace(sNameDevelopScript, ".js", "")
		End If	
    Else
        sNameDevelopScript = ""
        CommonScripts.Error "������ ���� ����� �������� ������ �� ��������� VBScript ��� JScript"
        Exit Sub
    End If
    iPos = InStrRev(sNameDevelopScript, "\")
    sNameDevelopScript = Mid(sNameDevelopScript, iPos + 1)

    If LCase(sNameDevelopScript) = "common" Then
        CommonScripts.Error "������ common.vbs ����� ������������� �������"
        Exit Sub
    End If

    ReloadScript
End Sub

' ������������� ������, ������� ����� �������������
' ����� �� ��� �������������
Sub SetupScriptForDevelopment()
    Dim strScriptPath

	strScriptPath = CommonScripts.SelectFileForRead(BinDir & "Config\Scripts\*.vbs", "VB �������|*.vbs|JS �������|*.js|��� �����|*")
               
	SetupScriptForDevelopmentInner strScriptPath

End Sub

' ���������/������������� ������ �:
' 1) ���� ������ � ������� ����, ����� �� ��������� ���� ������
' 2) ���� �������� ���������, ���������� ����� �� ������ ��������,
'����� ��������� ��������� ������
'
Sub ReloadScript()    
' TODO - ����� ��������, ����� ���� ������ ��� ��������� !
	If sDevelopScriptPath = "" Then 
		'��� �������, ����� ��� �� ������ ������� - �� ���� ������ ��� SetupScriptForDevelopment() ���������
		SetupScriptForDevelopment
		Exit Sub
	End If
    Set DevelopScript = CommonScripts.GetScriptByName(sNameDevelopScript)
    If DevelopScript Is Nothing Then
        RemoveVBAttribute sDevelopScriptPath
        Set DevelopScript = Scripts.Load(sNameDevelopScript)
        If DevelopScript Is Nothing Then
            CommonScripts.Error "�� ������� ��������� ������"
            Exit Sub
        End If
        iDevelopScriptIndex = CommonScripts.GetScriptIndexByName(sNameDevelopScript)
    Else
        RemoveVBAttribute sDevelopScriptPath
        iDevelopScriptIndex = CommonScripts.GetScriptIndexByName(sNameDevelopScript)
        Scripts.ReLoad iDevelopScriptIndex
    End If

    Set DevelopScript = CommonScripts.GetScriptByName(sNameDevelopScript)
    Set MacrosEnumerator = CreateObject("Macrosenum.Enumerator")
    MacrosArray = MacrosEnumerator.EnumMacros(DevelopScript)
    If UBound(MacrosArray) = -1 Then
        CommonScripts.Error "�������� � ������� �� �������"
        Exit Sub
    End If
    If UBound(MacrosArray) = 0 Then ' ����� ���� ������
        CommonScripts.CallMacros sNameDevelopScript, MacrosArray(0)
        Exit Sub
    End If

    ' ��� ���������� �������� ����� �� ������
    sMacrosList = Join(MacrosArray, vbCrLf) 
    sSelection = CommonScripts.SelectValue(sMacrosList, "�������� ������ ������")
    sSelection = CStr(sSelection)
    If sSelection = "" Then
        Exit Sub
    End If
    CommonScripts.CallMacros sNameDevelopScript, sSelection

End Sub

' ������ ��� ������ ������ ������� Attribute VB_Name = "..."
' ����������, ���� ������� ������ ����� Visual Basic
Sub RemoveVBAttribute(sFileName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sFileName) Then
        Exit Sub
    End If
    ForReading = 1
    Set f = fso.OpenTextFile(sFileName, ForReading, False)
    sTextOfFile = f.ReadAll
    f.Close

    arr = Split(sTextOfFile, vbCrLf)
    If InStr(Trim(arr(0)), "Attribute VB_Name") <> 1 Then
        Exit Sub
    End If

    arr(0) = "'" & Trim(arr(0))

    sTextOfFile = Join(arr, vbCrLf)
    ForWriting = 2
    Set f = fso.OpenTextFile(sFileName, ForWriting, False)
    f.Write sTextOfFile
    f.Close
End Sub

' ������ ������������ �������� ��������� ������� (��� ���, ��� ������ ������� � �������������)
Sub ReloadOpenedScript()
	set Doc = CommonScripts.GetTextDocIfOpened
	if Doc is nothing then Exit Sub

	Doc.Save
    
	SetupScriptForDevelopmentInner Doc.Path
	
End Sub
    
' ����� ������������� ���� � ����� �� ����� Config\Scripts
Function GetRelativePath(FilePath)
	GetRelativePath = ""
	set SrcFolder = CommonScripts.fso.GetFile(FilePath).ParentFolder 

	set Folder = SrcFolder
	i = 0
	do while (UCase(Folder.Name) <> "SCRIPTS") and (Folder.Name <> "")
		if i>5 then exit do
		set Folder = Folder.ParentFolder
	loop              
	if UCase(Folder.Name) <> "SCRIPTS" then Exit Function
	if UCase(Folder.ParentFolder.Name) <> "CONFIG" then Exit Function
			                                       
	sFolderPath = CommonScripts.fso.GetAbsolutePathName(Folder.Path)
	SrcFilePath = CommonScripts.fso.GetAbsolutePathName(FilePath)

	GetRelativePath = Replace(SrcFilePath, sFolderPath+"\", "")

End Function ' GetRelativePath
	
' ����� ������������� ���� � ����� � �����������
Function GetFilePathInRepository(FilePath)
	GetFilePathInRepository = ""

	sRepositoryPath1 = Trim(sRepositoryPath)
	if Right(sRepositoryPath1, 1) <> "\" then sRepositoryPath1 = sRepositoryPath1 + "\"
		
	if not CommonScripts.fso.FolderExists(sRepositoryPath1) then 
		Message "������� ����� ���� � �����������. ����� """ & sRepositoryPath & """ �� ����������", mRedErr
		Exit Function
	end if

	GetFilePathInRepository = sRepositoryPath1 + GetRelativePath(FilePath)

End Function ' GetFilePathInRepository
	
' ������ "SelectActionForScript" ������ ������ ���� ����������� ��������
'  	� ��� ���������� ������� ��������� ��������� ������������ �������� (������� �����, ������������� ������, ��������� ������)
'
Sub SelectActionForScript()
	set ScriptsDict = CreateObject("Scripting.Dictionary")
	                  
	' ���� � ������� ���� ����������� ������, ��� ������� ������ ���� � ������ ������ ������
	sOpenedScriptIndex = -1
	if not (Windows.ActiveWnd is Nothing) then
		sOpenedScript = UCase(CommonScripts.fso.GetAbsolutePathName(Windows.ActiveWnd.Document.Path))
		for i=0 to Scripts.Count-1
			if UCase(CommonScripts.fso.GetAbsolutePathName(Scripts.Path(i))) = sOpenedScript then 
				ScriptsDict.Add Scripts.Name(i), i
				sOpenedScriptIndex = i
				Exit for
			end if
		Next
	end if
	
	' ��� �������, ���������� � ������� ���, ����� ������ ���� � ������ ������ ������
	if iPrevSelectedScriptIndex <> - 1 then
		if iPrevSelectedScriptIndex <> sOpenedScriptIndex then
			ScriptsDict.Add Scripts.Name(iPrevSelectedScriptIndex), iPrevSelectedScriptIndex
		end if
	end if

	SelfScriptIndex = -1
	for i=0 to Scripts.Count-1
		if Scripts.Name(i) = SelfScript.Name then SelfScriptIndex = i ' ��� ����������� ������������ ����� �������

  		if (i <> sOpenedScriptIndex) and ( i <> iPrevSelectedScriptIndex ) then
			ScriptsDict.Add Scripts.Name(i), i
		end if
	Next
  
	ScriptIndex = CommonScripts.SelectValue(ScriptsDict)
	if ScriptIndex = "" then Exit Sub
		                                          
	' ��� ����������� ������������ ����� �������
	if UCase(Trim(ScriptIndex)) = UCase(Trim(SelfScript.Name)) then ScriptIndex = SelfScriptIndex
		
	iPrevSelectedScriptIndex = ScriptIndex ' �����������

	sScriptPath = Scripts.Path(ScriptIndex)
	sScriptName = """" & Scripts.Name(ScriptIndex) & """"

  	ActionList =	"������������� �����"+vbCrLf+ _
  					"�������������"+vbCrLf+ _
  					"���������"+vbCrLf+ _
  					"���������� � �����������"+vbCrLf+ _
  					"�������� �� �����������"+vbCrLf+ _
  					"�������� � ������� �� �����������"+vbCrLf+ _ 
					"������������� �� ������� ���������"
  					                        
	' ��� �������������� ������� ��������� ����������� ����������
  	if ScriptIndex = sOpenedScriptIndex then
	  	ActionList =	"�������������"+vbCrLf+ _
	  					"���������"+vbCrLf+ _
	  					"���������� � �����������"+vbCrLf+ _
	  					"�������� � ������� �� �����������"
  	end if

	Action = CommonScripts.SelectValue(ActionList)
	Select Case Action
		case "" Exit Sub

		case "������������� �����" 
			Documents.Open sScriptPath
			Scripts("ScriptMethodList").ShowMethodsList

		case "�������������" 
			Scripts.Reload(ScriptIndex)
			'Status "������ """ & Scripts.Name(ScriptIndex) & """ ������������"
			Status "������ " & sScriptName & " ������������"
			
		case "���������" 
			Scripts.Unload(ScriptIndex)
			Status "������ " & sScriptName & " ��������"

		case "���������� � �����������" 
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub

			CommonScripts.fso.CopyFile sScriptPath, sRepositoryFilePath, true
			Status "������ " & sScriptName & " ���������� � ����������� " & sRepositoryFilePath

		case "�������� �� �����������" 
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub
            
			sQuestion = "�� �������, ��� ������ �������� ������� ������ " & sScriptName & " ������� �� ����������� ?"
			if CommonScripts.wsh.Popup(sQuestion,, SelfScript.Name, 1+32) <> 1 then Exit Sub

			CommonScripts.fso.CopyFile sRepositoryFilePath, sScriptPath, true
			Status "������ " & sScriptName & " �������� �� ����������� " & sRepositoryFilePath
            
			Scripts.Reload(ScriptIndex)

		case "�������� � ������� �� �����������"
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub

			Scripts("�������� � ���������� ������").Diff2Files sRepositoryFilePath, sScriptPath

		case "������������� �� ������� ���������"
			EditFileWithExtEditorEx sScriptPath
			
		case Else
			Message "�������� """ & Action & """ ���� �� �����������", mRedErr
	End Select
	
End Sub ' SelectActionForScript

' ������� ������� � ������� ������ ����������� ��������
Sub SetNewRepositoryPath()
	Set ocReg = CommonScripts.Registry
	If ocReg Is Nothing Then 
		Message "�� ���� ������� ������ OpenConf.Registry", mRedErr
		Exit Sub
	End If
	Set SvcSvc = CreateObject("SvcSvc.Service")
	Path = SvcSvc.SelectFolder("������� ������� ����������� �������� (OpenConf_Scripts/�������)")
	If Path = "" Then
		Message "������� ����������� �������� �� �����", mInformation
		Exit Sub
	End If
	' �������� ��������� �������� � �������, ��� �� � ���������
	' ��� ������������ ��� �������������
	rk = ocReg.ScriptRootKey(SelfScript.Name)
	ocReg.Param(rk, "RepositoryPath") = Path
	sRepositoryPath = Path
	Message "������� ������� ����� �����������: " & sRepositoryPath
End Sub


' ��������������� ����������� � ������� 
' ���� � ������� ����� ����������� ��������
Private Sub InitRepositoryPath
	Set ocReg = CommonScripts.Registry
	If ocReg Is Nothing Then 
		Message "�� ���� ������� ������ OpenConf.Registry", mRedErr
		Message "���� � ����������� �������� ������ � ������� ������� (���������� sRepositoryPath)", mInformation
		Exit Sub
	End If
	rk = ocReg.ScriptRootKey(SelfScript.Name)
	Path = ocReg.Param(rk, "RepositoryPath")
	DoNotWarn = ocReg.Param(rk, "DoNotWarnAtStartUp")
	If IsNull(Path) or Path = "" Then
		' �������������� ������� ������ ��� ������ �������,	���� �� ���������� 	
		' ��� ������������� ��������, ��� ���� ������ �� ����������
		If IsNull(DoNotWarn) Then
			Message SelfScript.Name & ":: �� ����� ���� � ������� ������ ����������� ��������!", mRedErr
			Message "��� ������� ���� � ������� ����� ����������� �������� �������������� �������� " _
					& SelfScript.Name & "::InitRepositoryPath", mInformation
			ocReg.Param(rk, "DoNotWarnAtStartUp") = 1 ' ����� � ��������� ��� �� �������� ��� �������
		End If
		Exit Sub
	End If
	sRepositoryPath = Path
End Sub

' ������� ���� �� �������������� �� ������� ���������
' ������� ��� ������� ��������� �������� �� �������,
' ���� ������� ��� �� ������, ����������� ������ ��� ������
' ������������ ����� �������� ���������
' �� ��������� ���� ����� ������ � �������� (notepad.exe)
Sub EditFileWithExtEditorEx(Path)
	Dim EditorCmd, OCReg
	EditorCmd = ""
	Set ocReg = CommonScripts.Registry
	If Not ocReg Is Nothing Then
		RootKey = ocReg.ScriptRootKey(SelfScript.Name)
		EditorCmd = ocReg.Param(RootKey, "Editor")
		If IsNull(EditorCmd) or EditorCmd = "" Then
			EditorCmd = CommonScripts.SelectFileForRead("", "����������� ����� (*.exe)|*.exe")
			If EditorCmd <> "" Then
				ocReg.Param(RootKey, "Editor") = EditorCmd
			End If
		End If
	Else
		Message CommonScripts.GetLastError(), mRedErr		
	End If
	If EditorCmd = "" Then
		EditorCmd = "notepad.exe"
		Message "����� ����������� �������� �� ��������� (" & EditorCmd & ")", mInformation		
	End If
	CommonScripts.RunCommand EditorCmd, Path, 0
End	Sub

' ������ ��� �������� �������� ������������ ��� ����������
' ������� �� ������� ���������
Sub EditFileWithExtEditor()
	Dim Script	
	If sNameDevelopScript <> "" Then		
		ScriptIndex = CommonScripts.GetScriptIndexByName(sNameDevelopScript)	
		If ScriptIndex > -1 Then			
			EditFileWithExtEditorEx """" & Scripts.Path(ScriptIndex) & """"
		End If
	End If
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
	InitRepositoryPath
End Sub

Init 0 ' ��� �������� ������� ��������� �������������
