$NAME Разработка скриптов
' Для разработчиков скриптов
'
' Версия: $Revision: 1.14 $
'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057

' Для разработчиков скриптов имхо будет удобен скрипт "Разработка скриптов",
' который позволяет один раз задать отлаживаемый скрипт,
' и в дальнейшем через хоткей быстро перезагружать данный скрипт и вызывать его макросы.
' Если макросов несколько, сразу будет показан список для выбора нужного,
' если макрос всего один, он сразу же и выполнится.
' Еще фича: если создаем/редактируем скрипт как модуль в VB,
' из его текста убирается первая строчка "Attribute VB_Name",
' которую вставляет VB и на которую ругается Конфигуратор.
'
' макрос "SelectActionForScript" выдает список всех загруженных скриптов
'  	и для выбранного скрипта позволяет выполнить определенные действия (Открыть текст, перезагрузить скрипт, выгрузить скрипт)
'
Dim sRepositoryPath ' путь к репозитарию скриптов ( каждый разработчик устанавливает для себя свой путь )
	sRepositoryPath = "T:\Программирование\OpenConf\Репозитарий скриптов\OpenConf_Scripts\Скрипты" 

Dim sNameDevelopScript
Dim sDevelopScriptPath

Dim iPrevSelectedScriptIndex ' 
	iPrevSelectedScriptIndex = -1

' устанавливает скрипт, который будем разрабатывать
' сразу же его перезагружает
Sub SetupScriptForDevelopmentInner(sNameDevelopScriptParam)

	sNameDevelopScript = sNameDevelopScriptParam

    If sNameDevelopScript = "" Then
        Exit Sub
    End If

    sDevelopScriptPath = sNameDevelopScript
    If (InStrRev(UCase(sNameDevelopScript), ".VBS") > 0) Or (InStrRev(UCase(sNameDevelopScript), ".JS") > 0) Then

		' TODO $NAME не обязательно будет в первой строке, нужно перебирать строки до первой 
		' не пустой строки, плюс, игнорировать комментарии и $ENGINE
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
        CommonScripts.Error "Скрипт пока умеет работать только со скриптами VBScript или JScript"
        Exit Sub
    End If
    iPos = InStrRev(sNameDevelopScript, "\")
    sNameDevelopScript = Mid(sNameDevelopScript, iPos + 1)

    If LCase(sNameDevelopScript) = "common" Then
        CommonScripts.Error "Скрипт common.vbs нужно перезагружать вручную"
        Exit Sub
    End If

    ReloadScript
End Sub

' устанавливает скрипт, который будем разрабатывать
' сразу же его перезагружает
Sub SetupScriptForDevelopment()
    Dim strScriptPath

	strScriptPath = CommonScripts.SelectFileForRead(BinDir & "Config\Scripts\*.vbs", "VB скрипты|*.vbs|JS скрипты|*.js|Все файлы|*")
               
	SetupScriptForDevelopmentInner strScriptPath

End Sub

' загружает/перезагружает скрипт и:
' 1) если макрос в скрипте один, сразу же запускает этот макрос
' 2) если макросов несколько, предлагает выбор из списка макросов,
'затем запускает выбранный макрос
'
Sub ReloadScript()    
' TODO - нужна проверка, вдруг этот скрипт уже выгрузили !
	If sDevelopScriptPath = "" Then 
		'Так удобнее, когда его на хоткей вешаешь - не надо первый раз SetupScriptForDevelopment() запускать
		SetupScriptForDevelopment
		Exit Sub
	End If
    Set DevelopScript = CommonScripts.GetScriptByName(sNameDevelopScript)
    If DevelopScript Is Nothing Then
        RemoveVBAttribute sDevelopScriptPath
        Set DevelopScript = Scripts.Load(sNameDevelopScript)
        If DevelopScript Is Nothing Then
            CommonScripts.Error "Не удалось загрузить скрипт"
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
        CommonScripts.Error "Макросов в скрипте не найдено"
        Exit Sub
    End If
    If UBound(MacrosArray) = 0 Then ' всего один макрос
        CommonScripts.CallMacros sNameDevelopScript, MacrosArray(0)
        Exit Sub
    End If

    ' для нескольких макросов выбор из списка
    sMacrosList = Join(MacrosArray, vbCrLf) 
    sSelection = CommonScripts.SelectValue(sMacrosList, "Выберите нужный макрос")
    sSelection = CStr(sSelection)
    If sSelection = "" Then
        Exit Sub
    End If
    CommonScripts.CallMacros sNameDevelopScript, sSelection

End Sub

' убрать для модуля первую строчку Attribute VB_Name = "..."
' появляется, если создаем скрипт через Visual Basic
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

' макрос перезагрузки текущего открытого скрипта (для тех, кто правит скрипты в Конфигураторе)
Sub ReloadOpenedScript()
	set Doc = CommonScripts.GetTextDocIfOpened
	if Doc is nothing then Exit Sub

	Doc.Save
    
	SetupScriptForDevelopmentInner Doc.Path
	
End Sub
    
' Найти относительный путь к файлу от папки Config\Scripts
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
	
' Найти относительный путь к файлу в репозитарии
Function GetFilePathInRepository(FilePath)
	GetFilePathInRepository = ""

	sRepositoryPath1 = Trim(sRepositoryPath)
	if Right(sRepositoryPath1, 1) <> "\" then sRepositoryPath1 = sRepositoryPath1 + "\"
		
	if not CommonScripts.fso.FolderExists(sRepositoryPath1) then 
		Message "Неверно задан путь к репозитарию. Папка """ & sRepositoryPath & """ не существует", mRedErr
		Exit Function
	end if

	GetFilePathInRepository = sRepositoryPath1 + GetRelativePath(FilePath)

End Function ' GetFilePathInRepository
	
' макрос "SelectActionForScript" выдает список всех загруженных скриптов
'  	и для выбранного скрипта позволяет выполнить определенные действия (Открыть текст, перезагрузить скрипт, выгрузить скрипт)
'
Sub SelectActionForScript()
	set ScriptsDict = CreateObject("Scripting.Dictionary")
	                  
	' если в текущем окне редактируем скрипт, имя скрипта должно быть в начале списка выбора
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
	
	' имя скрипта, выбранного в прошлый раз, также должно быть в начале списка выбора
	if iPrevSelectedScriptIndex <> - 1 then
		if iPrevSelectedScriptIndex <> sOpenedScriptIndex then
			ScriptsDict.Add Scripts.Name(iPrevSelectedScriptIndex), iPrevSelectedScriptIndex
		end if
	end if

	SelfScriptIndex = -1
	for i=0 to Scripts.Count-1
		if Scripts.Name(i) = SelfScript.Name then SelfScriptIndex = i ' для возможности перезагрузки этого скрипта

  		if (i <> sOpenedScriptIndex) and ( i <> iPrevSelectedScriptIndex ) then
			ScriptsDict.Add Scripts.Name(i), i
		end if
	Next
  
	ScriptIndex = CommonScripts.SelectValue(ScriptsDict)
	if ScriptIndex = "" then Exit Sub
		                                          
	' для возможности перезагрузки этого скрипта
	if UCase(Trim(ScriptIndex)) = UCase(Trim(SelfScript.Name)) then ScriptIndex = SelfScriptIndex
		
	iPrevSelectedScriptIndex = ScriptIndex ' кеширование

	sScriptPath = Scripts.Path(ScriptIndex)
	sScriptName = """" & Scripts.Name(ScriptIndex) & """"

  	ActionList =	"Редактировать текст"+vbCrLf+ _
  					"Перезагрузить"+vbCrLf+ _
  					"Выгрузить"+vbCrLf+ _
  					"Копировать в репозитарий"+vbCrLf+ _
  					"Обновить из репозитария"+vbCrLf+ _
  					"Сравнить с версией из репозитария"+vbCrLf+ _ 
					"Редактировать во внешнем редакторе"
  					                        
	' для редактируемого скрипта некоторые возможности недоступны
  	if ScriptIndex = sOpenedScriptIndex then
	  	ActionList =	"Перезагрузить"+vbCrLf+ _
	  					"Выгрузить"+vbCrLf+ _
	  					"Копировать в репозитарий"+vbCrLf+ _
	  					"Сравнить с версией из репозитария"
  	end if

	Action = CommonScripts.SelectValue(ActionList)
	Select Case Action
		case "" Exit Sub

		case "Редактировать текст" 
			Documents.Open sScriptPath
			Scripts("ScriptMethodList").ShowMethodsList

		case "Перезагрузить" 
			Scripts.Reload(ScriptIndex)
			'Status "Скрипт """ & Scripts.Name(ScriptIndex) & """ перезагружен"
			Status "Скрипт " & sScriptName & " перезагружен"
			
		case "Выгрузить" 
			Scripts.Unload(ScriptIndex)
			Status "Скрипт " & sScriptName & " выгружен"

		case "Копировать в репозитарий" 
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub

			CommonScripts.fso.CopyFile sScriptPath, sRepositoryFilePath, true
			Status "Скрипт " & sScriptName & " скопирован в репозитарий " & sRepositoryFilePath

		case "Обновить из репозитария" 
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub
            
			sQuestion = "Вы уверены, что хотите заменить текущий скрипт " & sScriptName & " версией из репозитария ?"
			if CommonScripts.wsh.Popup(sQuestion,, SelfScript.Name, 1+32) <> 1 then Exit Sub

			CommonScripts.fso.CopyFile sRepositoryFilePath, sScriptPath, true
			Status "Скрипт " & sScriptName & " обновлен из репозитария " & sRepositoryFilePath
            
			Scripts.Reload(ScriptIndex)

		case "Сравнить с версией из репозитария"
			sRepositoryFilePath = GetFilePathInRepository(sScriptPath)
			if sRepositoryFilePath = "" then Exit Sub

			Scripts("Сравнить и объединить модуль").Diff2Files sRepositoryFilePath, sScriptPath

		case "Редактировать во внешнем редакторе"
			EditFileWithExtEditorEx sScriptPath
			
		case Else
			Message "Действие """ & Action & """ пока не реализовано", mRedErr
	End Select
	
End Sub ' SelectActionForScript

' Выбрать каталог с рабочей копией репозитория скриптов
Sub SetNewRepositoryPath()
	Set ocReg = CommonScripts.Registry
	If ocReg Is Nothing Then 
		Message "Не могу создать объект OpenConf.Registry", mRedErr
		Exit Sub
	End If
	Set SvcSvc = CreateObject("SvcSvc.Service")
	Path = SvcSvc.SelectFolder("Задайте каталог репозитория скриптов (OpenConf_Scripts/Скрипты)")
	If Path = "" Then
		Message "Каталог репозитория скриптов не задан", mInformation
		Exit Sub
	End If
	' Сохраним выбранное значение в реестре, что бы в следующий
	' раз использовать его автоматически
	rk = ocReg.ScriptRootKey(SelfScript.Name)
	ocReg.Param(rk, "RepositoryPath") = Path
	sRepositoryPath = Path
	Message "Выбрана рабочая копия репозитория: " & sRepositoryPath
End Sub


' Восстанавливает сохраненный в реестре 
' путь к рабочей копии репозитория скриптов
Private Sub InitRepositoryPath
	Set ocReg = CommonScripts.Registry
	If ocReg Is Nothing Then 
		Message "Не могу создать объект OpenConf.Registry", mRedErr
		Message "Путь к репозиторию придется задать в скрипте вручную (переменная sRepositoryPath)", mInformation
		Exit Sub
	End If
	rk = ocReg.ScriptRootKey(SelfScript.Name)
	Path = ocReg.Param(rk, "RepositoryPath")
	DoNotWarn = ocReg.Param(rk, "DoNotWarnAtStartUp")
	If IsNull(Path) or Path = "" Then
		' Предупреждение выводим только при первом запуске,	дабы не раздражало 	
		' тех пользователей скриптов, кто этот скрипт не использует
		If IsNull(DoNotWarn) Then
			Message SelfScript.Name & ":: Не задан путь с рабочей копией репозитория скриптов!", mRedErr
			Message "Для задания пути к рабочей копии репозитория скриптов воспользуйтесь макросом " _
					& SelfScript.Name & "::InitRepositoryPath", mInformation
			ocReg.Param(rk, "DoNotWarnAtStartUp") = 1 ' чтобы в следующий раз не ругаться при запуске
		End If
		Exit Sub
	End If
	sRepositoryPath = Path
End Sub

' Открыть файл на редактирование во внешнем редакторе
' Команда для запуска редактора читается из реестра,
' если команда еще не задана, открывается диалог для выбора
' исполняемого файла внешнего редактора
' По умолчанию файл будет открыт в Блокноте (notepad.exe)
Sub EditFileWithExtEditorEx(Path)
	Dim EditorCmd, OCReg
	EditorCmd = ""
	Set ocReg = CommonScripts.Registry
	If Not ocReg Is Nothing Then
		RootKey = ocReg.ScriptRootKey(SelfScript.Name)
		EditorCmd = ocReg.Param(RootKey, "Editor")
		If IsNull(EditorCmd) or EditorCmd = "" Then
			EditorCmd = CommonScripts.SelectFileForRead("", "Исполняемые файлы (*.exe)|*.exe")
			If EditorCmd <> "" Then
				ocReg.Param(RootKey, "Editor") = EditorCmd
			End If
		End If
	Else
		Message CommonScripts.GetLastError(), mRedErr		
	End If
	If EditorCmd = "" Then
		EditorCmd = "notepad.exe"
		Message "Будет использован редактор по умолчанию (" & EditorCmd & ")", mInformation		
	End If
	CommonScripts.RunCommand EditorCmd, Path, 0
End	Sub

' Макрос для открытия текущего назначенного для разработки
' скрипта во внешнем редакторе
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
' Процедура инициализации скрипта
'
Sub Init(dummy) ' Фиктивный параметр, чтобы процедура не попадала в макросы
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "Не могу создать объект OpenConf.CommonServices", mRedErr
        Message "Скрипт " & SelfScript.Name & " не загружен", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	SelfScript.AddNamedItem "CommonScripts", c, False
	InitRepositoryPath
End Sub

Init 0 ' При загрузке скрипта выполняем инициализацию
