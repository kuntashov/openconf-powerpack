$NAME Сравнить объект с предыдущей версией

'    Сравнение объектов конфигурации при помощи gcomp
'
' Сравнивает текущий объект (справочник, документ, отчёт) с тем, что лежит в каталоге Src.
' Для сравнения выполняется выборочная декомпиляция текущего объекта в каталог Src_New,
' и затем запускается <DiffCommand> Src Src_New
'

Dim DiffCommand
DiffCommand = "kdiff3"


Private Function GetObjPath(FilterStr)
	Dim aType, aKind
	
	GetObjPath = ""
	FilterStr = ""

	Set CommonScripts = CreateObject("OpenConf.CommonServices")
	Set Doc = Windows.ActiveWnd.Document
	'Message Doc.Name, mNone

	Set RE = New RegExp
	
	'форма объекта    Справочник.Номенклатура.Форма
	RE.Pattern = "([^\s.]+)\.([^\s.]+)\..*"
	match = RE.Test(Doc.Name)
	If not match Then
		'диалог редактирования свойств объекта    CMDSubDoc::Справочник Номенклатура
		RE.Pattern = "CMD[a-zA-Z]+Doc::(\S+)\s+(\S+)"
		match = RE.Test(Doc.Name)
	End If
	If not match Then
		'конфигурация CMDTabDoc::Конфигурация
		RE.Pattern = "CMD[a-zA-Z]+Doc::(\S+)(.*)"
		match = RE.Test(Doc.Name)
		If match Then
			'Message "MDWnd = " & MDWnd.GetSelected
			parts = split(MDWnd.GetSelected, "\")
			'Message "UBound = " & UBound(parts)
			If parts(0) = "" Then Exit Function
			GetObjPath = parts(0)
			If parts(0) = "Планы Счетов" Then
				GetObjPath = "ПараметрыБухгалтерии"
			ElseIf parts(0) = "Виды Субконто" Then
				GetObjPath = "ПараметрыБухгалтерии"
			ElseIf parts(0) = "Операция" Then
				GetObjPath = "ПараметрыБухгалтерии"
			ElseIf parts(0) = "Проводка" Then
				GetObjPath = "ПараметрыБухгалтерии"
			ElseIf parts(0) = "Константы" Then
				GetObjPath = parts(0)
			ElseIf parts(0) = "Перечисления" Then
				GetObjPath = parts(0)
			ElseIf parts(0) = "Регистры" Then
				GetObjPath = parts(0)
			ElseIf UBound(parts) >= 1 Then
				GetObjPath = GetObjPath & "\" & parts(1)
			End If
			FilterStr = " --filter " & GetObjPath
			GetObjPath = "\" & GetObjPath
			'Message "FilterStr = " & FilterStr & ", " & UBound(parts)
			Exit Function
		End If
	End If
	If not match Then Exit Function
	
	Set Matches = RE.Execute(Doc.Name)
	aType = LCase(Matches(0).Submatches(0))
	aKind = Matches(0).Submatches(1)

	Select Case aType
		Case "справочник" aType = "Справочники"
		Case "документ"   aType = "Документы"
		Case "регистр"    aType = "Регистры"
			aKind = ""
		Case "отчет"      aType = "Отчеты"
		Case "обработка"  aType = "Обработки"
		Case "конфигурация"  aType = ""
		Case Else Exit Function
	End Select
	
	If aType <> "" Then
		If aKind <> "" Then
			GetObjPath = "\" & aType & "\" & aKind
			FilterStr = " --filter " & aType & "." & aKind
		Else
			GetObjPath = "\" & aType
			FilterStr = " --filter " & aType
		End If
	End If
End Function

Sub Decompile(Dir)
	Dim ObjPath, FilterStr
	
	ObjPath = GetObjPath(FilterStr)

	Set CommonScripts = CreateObject("OpenConf.CommonServices")
	
	FullDirName = """" & IBDir & Dir & """"
	Arguments = " -d -vv -F """ & IBDir & "1cv7.md"" -D " & FullDirName & FilterStr
	CommonScripts.RunCommandAndWait "gcomp", Arguments
	'Message Arguments, mNone
End Sub


Sub DiffCurrentObject()
	Dim ObjPath, FilterStr
	Set CommonScripts = CreateObject("OpenConf.CommonServices")

	Decompile "Src_New"

	ObjPath = GetObjPath(FilterStr)
	Src = """" & IBDir & "Src" & ObjPath & """"
	Src_New = """" & IBDir & "Src_New" & ObjPath & """"
	Arguments = Src & " " & Src_New
	CommonScripts.RunCommand DiffCommand, Arguments, false
End Sub

Sub DecompileCurrentObject()
	Decompile "Src"
End Sub

