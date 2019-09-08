'      Создать процедуру и кнопку для нее на форме
'
'         Мой e-mail: artbear@bashnet.ru
'         Мой ICQ: 265666057
'
' следующие параметры можно поменять, чтобы изменить работу скрипта
Dim gsScriptCreateButtonName
gsScriptCreateButtonName = "Создать кнопку на форме"

Dim gsProceduresSignature
gsProceduresSignature = "!ИмяПроцедурыВместеСКнопкой!"
Dim gsFunctionSignature
gsFunctionSignature = "!ИмяФункцииВместеСКнопкой!"

' следующие две строки только для вывода отладочных сообщений
' их можно удалить или закомментировать
Dim DebugFlag 'обязательно глобальная переменная
'DebugFlag = True 'Разрешаю вывод отладочных сообщений

' Обработка события "Вставка шаблона".
' Возникает при вставке шаблона перед обработкой его 1С.
' Позволяет изменить текст замены, скорректировав его по ситуации.
' name - имя вставляемого шаблона.
' text - текст замены шаблона. Можно изменить его.
' cancel - флаг отмены. При установке а True вставка шаблона отменяется.
'
Sub Telepat_OnTemplate(Name, Text, Cancel)
		if InStr(Text, gsProceduresSignature) Then
    	Cancel = CreateProceduresAndHisButtonLocal(False, Text)
    Elseif InStr(Text, gsFunctionSignature) then
    	Cancel = CreateProceduresAndHisButtonLocal(True, Text)
    End If
End Sub

Function CreateProceduresAndHisButtonLocal(bIsFunction, sText)
	CreateProceduresAndHisButtonLocal = True      ' По умолчанию отменим вставку шаблона

'	DebugEcho "sText",  sText
	if bIsFunction then
		sType = "функции"
	else
		sType = "процедуры"
	end if

	' Дадим пользователю выбрать, как назвать процедуру/функцию
	if not DebugFlag then
		sProcName = InputBox("Укажите имя "+sType, "Ввод "+sType+" и кнопки для нее")
	else
		sProcName = "Выполнить"
	end if
	If Len(sProcName) = 0 Then Exit Function

	' Подставим в шаблон выбранное имя переменной
	if bIsFunction then
		sText = Replace(sText, gsFunctionSignature, sProcName, 1, -1, 1)
	else
		sText = Replace(sText, gsProceduresSignature, sProcName, 1, -1, 1)
	end if
	DebugEcho "sText",  sText

	set ScriptCreateButton = GetScriptByName(gsScriptCreateButtonName)
	if IsEmpty(ScriptCreateButton) then
		exit Function
	End if

	ScriptCreateButton.CreateButton "Закрыть", sProcName, sProcName+"()", "Кнопка_"+sProcName

	CreateProceduresAndHisButtonLocal = False ' Не будем отменять вставку шаблона

End Function ' CreateProceduresAndHisButtonLocal

' вставить в текущую позицию курсора замененный шаблон
Sub CreateCodeAndHisButtonFromText(bIsFunction, sText)
	set doc = GetTextDoc(null)
	if IsNull(doc) then Exit Sub

	Telepat.ConvertTemplate(sText)

	Cancel = CreateProceduresAndHisButtonLocal(bIsFunction, sText)

	if Cancel then Exit Sub

  ' установить позицию курсора
  bFindPosition = false
  MyArray = Split(sText, vbCRLF, -1, 1)
  for i=LBound(MyArray) to UBound(MyArray)
  	Pos = inStr(MyArray(i), "<?>")
  	if Pos > 0 then
		  bFindPosition = True
  		exit for
  	end if
  next
  sText = Replace(sText, "<?>", "", 1, -1, 1)

  PosLine = doc.SelStartLine + i
  PosCol = doc.SelStartCol + Pos-1

 	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = sText

  if bFindPosition then
  	doc.MoveCaret PosLine, PosCol
  end if

End Sub ' CreateCodeAndHisButtonFromText

' это макрос для интерактивного создания процедуры и кнопки
Sub CreateProceduresAndHisButton()
	sText = "//----------------------------- -----------------------"
	sText = sText + vbCrLF + "// "+gsProceduresSignature+"()"
	sText = sText + vbCrLF + "Процедура "+gsProceduresSignature+"() Экспорт"
	sText = sText + vbCrLF + vbTab +"<?>"
	sText = sText + vbCrLF + "КонецПроцедуры // "+gsProceduresSignature+ vbCrLF

	CreateCodeAndHisButtonFromText False, sText
End Sub ' CreateProceduresAndHisButton

' это макрос для интерактивного создания процедуры и кнопки
Sub CreateFunctionAndHisButton()
	sText = "//----------------------------- -----------------------"
	sText = sText + vbCrLF + "// "+gsFunctionSignature+"()"
	sText = sText + vbCrLF + "Функция "+gsFunctionSignature+"() Экспорт"
	sText = sText + vbCrLF + vbTab +"<?>"
	sText = sText + vbCrLF + "КонецФункции // "+gsFunctionSignature+ vbCrLF

	CreateCodeAndHisButtonFromText True, sText
End Sub ' CreateFunctionAndHisButton

' -----------------------------------------------------------------------------
' получить скрипт по названию
' если скрипт отсутствует, возвращается empty
'
' почему-то конструкция
'          set GetScript = Scripts(ScriptName)
' не хочет работать,  если ScriptName - обычная строковая переменная
'
Function GetScriptByName(ScriptName)
	GetScript = Empty

	bFindScript = False
	for i = 0 to Scripts.Count-1
		if LCase(Scripts.Name(i)) = LCase(ScriptName) then
			bFindScript = True
			exit for
		end if
	next
	if not bFindScript then
		MsgBox("Скрипт для добавления кнопки не найден! " +vbCRLF+vbCRLF+"Скрипт "+""""+sScriptCreateButtonName+".vbs"+""""+" отсутствует!")
		exit Function
	End if

	DebugEcho "Name",  Scripts.Name(i)

	set GetScriptByName = Scripts.Item(i)
End Function ' GetScriptByName

' пустой параметр, чтобы не попадал в макросы
Function GetTextDoc(param)
	GetTextDoc = null

  If Windows.ActiveWnd Is Nothing Then
     MsgBox "Нет активного окна"
     Exit Function
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     MsgBox "Окно не текстовое"
     Exit Function
  End If

	set GetTextDoc = doc
End Function ' GetTextDoc

Sub Echo(text)
	Message text, mNone
End Sub'Echo

Sub Debug(title, msg)
on error resume next
  DebugFlag = DebugFlag
  if err.Number<>0 then
    err.Clear()
    on error goto 0
    Exit Sub
  end if
  if DebugFlag then
    if not (IsEmpty(msg) or IsNull(msg)) then
      msg = CStr(msg)
    end if
    if not (IsEmpty(title) or IsNull(title)) then
      title = CStr(title)
    end if
    If msg="" Then
      Echo(title)
    else
      Echo(title+" - "+msg)
    End If
  End If
on error goto 0
End Sub'Debug

Sub DebugEcho(title, msg)
	Debug title, msg
End Sub'DebugEcho

' Инициализация скрипта. param - пустой параметр,
' чтобы не попадал в макросы
'
Sub Init(param)
    Set t = Plugins("Телепат")  ' Получаем плагин
    If Not t Is Nothing Then    ' Если "Телепат" загружен
        ' Привязываем скрипт к событиям плагина
        SelfScript.AddNamedItem "Telepat", t, False
    Else
        ' Плагин не загружен. Выгрузим и скрипт
        Scripts.Unload SelfScript.Name
    End If
End Sub

' При загрузке скрипта инициализируем его
Init 0
