$NAME Меню макросов из файла
'      Скрипт для добавления макросов из скриптов в меню модуля по Ctrl-2 (c) artbear,  2004
'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057

' Скрипт работает в паре с плагином "Телепат"
' и обрабатывает события от него
' разместите в bin\config\scripts
'
' Для нормальной работы в каталоге bin\config должен быть файл Macros.ini
' c определением макроса из скрипта
' Формат файла:
' ИмяСкрипта.ИмяМакроса = здесь наименование макроса, которое появится в меню
' можно использовать символ "-" на отдельной строке,  это будет разделитель меню
' если указать [-],  тогда будет разделитель групп макросов
'
' группы задаются в квадратных скобках [Настройки],  [Работа с текстом]
'
' Естественно, что макросы можно описывать в нескольких группах
' Например,
' 1C++.OpenDefClsPrm = Открыть файл определений 1С++ (defcls.prm)
'
' Посмотрите приложенный файл Macros.ini - там все есть
'
Dim gsMacrosFileName ' расположен по умолчанию в bin\config
  gsMacrosFileName = "config\Macros.ini"

  sConstMenuPrefix = "Меню макросов из файла"

Dim Telepat

'  Dim ResDict 'as Dictionary
Dim MenuForTelepat ' as string

' Обработчик события "Вызов меню шаблонов"
' Позволяет динамически добавить команды в меню шаблонов.
' Для этого должен вернуть строку-описатель добавляемых пунктов.
' Каждый добавляемый пункт должен располагаться на отдельной строке.
' Вложенные группы команд определяются по табуляции в начале строки.
' После названия команды через символ | можно указать символы
' Далее через символ | можно указать идентификатор команды.
' В этом случае в событие OnCustomMenu вместо названия пункта меню
' будет передан этот идентификатор
' Для создания разделителей укажите имя "-"
'
Function Telepat_GetMenu()   
  'Telepat_GetMenu = Telepat_GetMenu & MenuForTelepat
  RefreshMacros
  Telepat_GetMenu = MenuForTelepat
End Function ' Telepat_GetMenu

' Обработчик события OnCustomMenu.
' Вызывается при выборе пользователем пункта меню,
' добавленного в "GetMenu"
' Cmd - название (или идентификатор) выбранного пункта меню.
'
Sub Telepat_OnCustomMenu(CmdArtur)
	Dim M
	s = Trim(CStr(CmdArtur))
	If (InStr(s, sConstMenuPrefix) = 1) Then
		On Error Resume Next
		' Используем универсальный способ для вызова макросов из скриптов
		' (Вызов с использованием Eval не учитывал особенностей js: то, что в
		' js нет процедур и возможность задания идентификаторов на русском языке -- a13x
		MacrosSpec = Replace(CmdArtur, sConstMenuPrefix, "")
		M = Split(MacrosSpec, "::", 2)
		CommonScripts.CallMacros M(0), M(1)
		On Error GoTo 0
	End If
End Sub ' Telepat_OnCustomMenu

' получить данные из INI-файла
Function GetDataFromIniFile(ByVal IniFileName, ByVal FlagInsertMenuGroup)
	GetDataFromIniFile = False

	' далее автоматически
	Dim IniFile 'As TextStream

	On Error Resume Next
	Dim ForRead
	ForRead = 1

	Dim fso 'as FileSystemObject
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set IniFile = fso.OpenTextFile(IniFileName, ForRead)
	If Err.Number <> 0 Then
		Err.Clear()
		CommonScripts.Echo "Ini-файл " & IniFileName & " не удалось открыть!"
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
		MenuForTelepat = "Макросы с удобным названием" & vbCrLf
	Else
		MenuForTelepat = ""
	End If
	TabForGroup = ""

	bFind = False
	Do While IniFile.AtEndOfStream <> True
		s = Trim(IniFile.ReadLine)
	' если не строка-комментарий
		If Not RegExpTest("^\s*[;']", s) Then
		If Left(s, 1) = "-" Then ' разделитель
'        If RegExpTest("^\s*\[-\]\s*", s) Then ' разделитель
		  MenuForTelepat = MenuForTelepat & TabForGroup & vbTab & "-" & vbCrLf
		Else
		Set Matches = regForGroup.Execute(s) ' это группа меню
		If Matches.Count > 0 Then
	'        CommonScripts.Echo "Группа " & Matches(0).SubMatches(0)
		    If RegExpTest("^\s*\[-\]\s*", s) Then ' разделитель
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
	' выделить ключ и значение в Ini-файле, убрав возможные комментарии
		  Set Matches = reg.Execute(s)
		  If Matches.Count > 0 Then
	' добавить новую пару, исключив из значения табуляцию и левые(и правые) пробелы
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

	'if ResDict.Count=0 then ' не нашли ни одного ключа
	If bFind = False Then ' не нашли ни одного ключа
		CommonScripts.Echo "Не удалось прочесть данные из Ini-файла " & IniFileName
		GetDataFromIniFile = False
	Else
		GetDataFromIniFile = True
	End If
End Function ' GetDataFromIniFile

' проверить на соответствие шаблону
' регистр символов не важен
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

' записать все макросы в файл "config\Macros_all.ini",
' где их можно приготовить к использованию в меню
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
		arr = e.EnumMacros(script) ' Получение массива макросов объекта
		For j = 0 To UBound(arr)
			IniFile2.WriteLine sScriptName & "." & arr(j) & " ="
		Next
	Next

End Function ' SaveAllMacrosToFile

' чтобы в список попали все скрипты, иначе только те,
' что успели загрузиться до загрузки этого скрипта
Sub Configurator_AllPluginsInit()
	'Init 0
	RefreshMacros
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
	
	' При загрузке скрипта инициализируем его
  	c.AddPluginToScript SelfScript, "Телепат", "Telepat", Telepat
  
	SelfScript.AddNamedItem "CommonScripts", c, False

	RefreshMacros
End Sub

Init 0

Sub ShowStructuredMenuOfMacros()
	GetDataFromIniFile BinDir & gsMacrosFileName, False	
	Set srv = CreateObject("Svcsvc.Service")
	Telepat_OnCustomMenu srv.PopupMenu(MenuForTelepat, 0)
End Sub
