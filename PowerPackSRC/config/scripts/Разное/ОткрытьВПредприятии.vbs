'   Макрос для открывания объектов в режиме предприятия.
'   
'Находясь в форме объекта (в конфигураторе) запускаем метод Open(), и текущий 
'объект будет открыт в предприятии. Если предприятие было запущено из конфигуратора
'(по F11 или скриптом), то объект будет открыт в нём без вопросов.
'
'Открывать можно:
'  - внешние отчёты
'  - текстовые файлы
'  - отчёты и обработки конфигурации
'  - справочники (будет открыта форма списка)
'  - журналы документов
'  
'Для работы требуется программа OpenIn1C.exe. Её нужно положить в каталог Bin, 
'либо в Bin\Config\System, либо в каталог, прописанный в PATH
'
' Автор: Алексей Диркс aka ADirks
' e-mail: <adirks@ngs.ru>


Sub Open()
	set ob = CreateObject("1S.IBState")
	iIBState = ob.IBState(IBDir)
	If iIBState = -1 Then
		Set RunScript = Scripts("Разбор командной строки Конфигуратора")
		Scripts("Шорткаты").Run1CInExlusiveMode()
	ElseIf iIBState = 0 Then
		Set RunScript = Scripts("Разбор командной строки Конфигуратора")
		RunScript.RunInSharedMode()
	End If

	Set Doc = CommonScripts.GetTextDoc(true, 0)
	If Doc Is Nothing Then Exit Sub
	'Message Doc.Name
	'Message Doc.Path
	
	fname = BinDir & "OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\System\OpenIn1C.exe"
	If not CommonScripts.FSO.FileExists(fname) Then fname = BinDir & "Config\Data\OpenIn1C\OpenIn1C.exe" ' имхо, т.к у меня нет каталога Config\Sys - пусть все будет в одном месте
	If not CommonScripts.FSO.FileExists(fname) Then fname = "OpenIn1C.exe"


	
	PauseTime = 2000 'in milliseconds
	TypeLetter = ""
	if (Doc.ID >= 0) then ' не внешний файл
		Set MetaObj = GetMetaObject(Doc.Name, TypeLetter)
		If Not MetaObj Is Nothing And TypeLetter <> "" Then
			parts = split(Doc.Name, ".")
			Repr = MetaObj.Name
			If MetaObj.Present <> "" Then Repr = MetaObj.Present
			If MetaObj.Descr <> "" Then Repr = Repr & " (" & MetaObj.Descr & ")"
			
			If parts(0) = "Документ" Then Repr = "(Полный)" 'Для документов будем просто открывать полный журеал
			
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
	If parts(0) = "Справочник" Then TypeLetter = "-s"
	If parts(0) = "Журнал" Then TypeLetter = "-j"
	If parts(0) = "Документ" Then TypeLetter = "-j"
	If parts(0) = "Отчет" Then TypeLetter = "-r"
	If parts(0) = "Обработка" Then TypeLetter = "-p"
End Function


'======= Инициализация =============================
Private Sub Init()
	Set c = Nothing
	On Error Resume Next
	Set c = CreateObject("OpenConf.CommonServices")
	On Error GoTo 0
	If c Is Nothing Then
		Message "Не могу создать объект OpenConf.CommonServices", mRedErr
		Message "Скрипт " & SelfScript.Name & " не загружен", mInformation
		Scripts.UnLoad SelfScript.Name
	Else
		c.SetConfig Configurator
		SelfScript.AddNamedItem "CommonScripts", c, False
	End If
End Sub

Init
'===================================================
