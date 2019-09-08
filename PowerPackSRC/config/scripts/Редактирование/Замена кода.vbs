' Макросы для замены и/или обрамления кода
'
'	CodeReplace
'	CodeReplaceWithQuestion
'	CodeFramingWithQuestion
'	CodeFramingForTimer
'
' Например, если стоим в строке
'        Если Код <> КодИзСправочника Тогда
' и запускаем макрос, эта строка заменится на
'
' // -- Артур --
' //        Если Код <> КодИзСправочника Тогда
'         Если Код <> КодИзСправочника Тогда
' // -- завершение
'
' Макрос "Копировать строку (или несколько выделенных строк"
'	CopyLine
'
' Макрос "обменять местами левую и правую часть присвоения с сохранением возможных комментариев и пробелов возле знака присвоения"
'
'	ExchangeLeftAndRightOfAssign
'
' например, вместо
' 		КолСубс1 =  гМассивПараметровДляРасчета[ПРМ_КолСубс1];  // пример 
' будет
' 		гМассивПараметровДляРасчета[ПРМ_КолСубс1] =  КолСубс1;  // пример 
'
' Макросы "комментировать текст" и "раскомментировать текст"
'
'	CommentSelection
'	UnCommentSelection
'
' в отличие от типового ставят/убирают комментарии, 
' даже если ничего не выделено или выделен текст на одной строке
' правильно комментирует vbs-скрипты (символом апострофа)
'
' ------------------------------------------------------------------------------------------------
' есть код для облегчения работы с Телепатом.
'
' Принцип работы:
' когда в конструкции типа 
'      << Перем1 = |Перем2; >>
' где | - положение курсора,
' 
' начинаем набирать первые буквы какой-нибудь функции, например, "Сокр",
' Телепата выводит список возможных вариантов (СокрЛ, СокрЛП, СокрП),
' после выбора, например, СокрЛП, исходный код преобразуется в 
'     << Перем1 = СокрЛП()Перем2; >> // смотрится коряво, и требуется вручную расставлять скобки
'                               
' что выглядит не очень красиво.
' Так вот следующий код позволяет при тех же самых действиях из строки
'     << Перем1 = |Перем2; >>
' получить строку
'     << Перем1 = СокрЛП(Перем2);  >> // выглядит намного лучше, не правда ли
' 
' т.е. то, что в большинстве случае мы и хотим получить 
' 
' PS в следующей версии попытаюсь получить код, который будет позволять отмену
' т.е. после получения << Перем1 = СокрЛП(Перем2); >>, нажав Ctrl-Z (или Действия - Отменить),
' можно будет получить строку << Перем1 = СокрЛП()Перем2); >>
' и таким образом, пользователю легко выбирать нужное ему поведение при вставке функции
' ------------------------------------------------------------------------------------------------
'
'
' скрипт умеет работать с несколькими выделенными строками, 
' строки не обязательно должны быть выделены полностью
' если нет выделения, скрипт работает с текущей строкой
'  
' строки начальных и конечных комментариев поменять очень легко, 
' они находятся на пару строк ниже
'
' Автор:	Артур Аюханов aka artbear
' e-mail:	artbear@bashnet.ru
'

' Если true, тогда "Перем1 = СокрЛП|Перем2;" меняется на "Перем1 = СокрЛП(Перем2)|;"
' Если false, тогда "Перем1 = СокрЛП|Перем2;" меняется на "Перем1 = СокрЛП(|Перем2;" - говорят, что аля v8
bDontInsertEndCommaInTelepatReplaceFlag = true

sBeginComment = vbCrLf & "// -- Артур -- " ' Меняем здесь -  НАЧАЛЬНЫЙ коммент
sEndComment = "// -- завершение" & vbCrLf ' Меняем здесь -   ЗАВЕРШАЮЩИЙ коммент

' Макросы для замены и/или обрамления кода
'
' Например, если стоим в строке
'        Если Код <> КодИзСправочника Тогда
' и запускаем макрос, эта строка заменится на
'
' // -- Артур --
' //        Если Код <> КодИзСправочника Тогда
'         Если Код <> КодИзСправочника Тогда
' // -- завершение
'  
' строки начальных и конечных комментариев поменять очень легко, 
' они находятся на пару строк выше (sBeginComment и sEndComment)
'
Sub CodeReplace()
	CodeReplaceInner sBeginComment, sEndComment, true
End Sub ' CodeReplace 
          
' то же, что предыдущий, но оба выражения вводится пользователем
Sub CodeReplaceWithQuestion()
	sFrameBegin = InputBox("Укажите начало обрамляющего блока", SelfScript.Name)
	sFrameEnd = InputBox("Укажите завершение обрамляющего блока", SelfScript.Name)
	CodeReplaceInner sFrameBegin, sFrameEnd, true
End Sub ' CodeReplaceWithQuestion

' Обрамить строку (или выделенные строки)
' некоторым текстом.
' Например: вместо строки 
'	Результат = Выполнить();
' можно получить 
'	глЗапускТаймера("Выполнить");
'	Результат = Выполнить();
'	глОстановТаймера("Выполнить");
'
Sub CodeFramingInner(FrameBegin, FrameEnd)
	CodeReplaceInner FrameBegin, FrameEnd, false
End Sub ' CodeFramingInner 

' Макрос для вызова CodeFramingInner
' Обрамить строку (или выделенные строки)
' некоторым текстом.
' Например: вместо строки 
'	Результат = Выполнить();
' можно получить 
'	глЗапускТаймера("Выполнить");
'	Результат = Выполнить();
'	глОстановТаймера("Выполнить");
'
Sub CodeFramingWithQuestion()
	sFrameBegin = InputBox("Укажите начало обрамляющего блока", SelfScript.Name)
	sFrameEnd = InputBox("Укажите завершение обрамляющего блока", SelfScript.Name)
	CodeFramingInner sFrameBegin, sFrameEnd
End Sub ' CodeFramingWithQuestion

' Обрамить строку (или выделенные строки)
' текстом для тайминга
' Например:
'	гТаймер.Запуск("Выполнить");
'	Результат = Выполнить();
'	гТаймер.Стоп("Выполнить");
Sub CodeFramingForTimer()
	sText = InputBox("Укажите название метода для таймера", SelfScript.Name)
	CodeFramingInner "гТаймер.Запуск(""" & sText & """);", "гТаймер.Стоп(""" & sText & """);"
End Sub ' CodeFramingForTimer

' Метод для обрамления некоторого кода (выделенных строк или одной строки, если ничего не выделено)
' специальными строками
' старый код может комментироваться
'
' метод умеет работать с несколькими выделенными строками, 
' строки не обязательно должны быть выделены полностью
' если нет выделения, скрипт работает с текущей строкой
'
Sub CodeReplaceInner(FrameBegin, FrameEnd, bDuplicateBeforeAsComment)
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub

    Line1 = doc.SelStartLine
    Line2 = doc.SelStartLine
    Col1 = 0
    Col2 = doc.LineLen(doc.SelStartLine)

    If (doc.SelStartLine <> doc.SelEndLine) Then ' есть выделение на нескольких строках
        If (doc.SelStartCol = 0) And (doc.SelEndCol = 0) Then ' выделено ровно одна или несколько строк
            Line2 = doc.SelEndLine - 1
            Col2 = doc.LineLen(doc.SelEndLine - 1)
        Else
            Line2 = doc.SelEndLine
            Col2 = doc.LineLen(doc.SelEndLine)
        End If
    End If
    CurrentText = doc.Range(Line1, Col1, Line2, Col2)

    sArray = Split(CurrentText, vbCrLf)
    CurrentText = FrameBegin & vbCrLf
	if bDuplicateBeforeAsComment then
	    For Each s In sArray
	        CurrentText = CurrentText & "//" & s & vbCrLf
	    Next
	end if
	
    For Each s In sArray
        CurrentText = CurrentText & s & vbCrLf
    Next
    CurrentText = CurrentText & FrameEnd

    doc.Range(Line1, Col1, Line2, doc.LineLen(Line2)) = CurrentText

    ' верну курсор в ту же позицию с учетом добавленных строк
	if bDuplicateBeforeAsComment then
	    iNewLine = Line1 + UBound(Split(FrameBegin, vbCrLf)) + UBound(sArray) + 2
	else
	    iNewLine = Line1 + UBound(Split(FrameBegin, vbCrLf)) + 1
	end if
	
    doc.MoveCaret iNewLine, Col1, iNewLine, Col1

End Sub ' CodeReplaceInner 

' Макрос "Копировать строку (или несколько выделенных строк"
' работает именно для строк, положение курсора неважно, 
' может быть выделено несколько строк или ничего не выделено, 
' выделение может быть неполным (т.е. не вся строка), 
' после дублирования строки.строк в новой строке курсор встает в ту же колонку, в которой стоял до запуска макроса
'
Sub CopyLine()
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub 
		
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = doc.Range(doc.SelStartLine) & vbCrLf & doc.Range(doc.SelStartLine)
  	else 
		if doc.SelEndLine < doc.LineCount-1 then
    		doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) = doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) & doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0)
		else
    		doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine)) = doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))  & vbCrLf & doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))
		end if
  	end if

    ' верну курсор в ту же позицию  на новой строке с учетом добавленных строк
    doc.MoveCaret doc.SelEndLine + 1, doc.SelStartCol, doc.SelEndLine + 1, doc.SelStartCol
End Sub ' CopyLine

' еще один вариант (потолще)
Sub CopyLine0()
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub
  
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = doc.Range(doc.SelStartLine) & vbCrLf & doc.Range(doc.SelStartLine)
	    ' верну курсор в ту же позицию с учетом добавленных строк
	    doc.MoveCaret doc.SelStartLine + 1, doc.SelStartCol, doc.SelStartLine + 1, doc.SelStartCol
		
  	else  ' есть выделение на нескольких строках
	    Line1 = doc.SelStartLine
	    Line2 = doc.SelStartLine
	    Col1 = 0
	    Col2 = doc.LineLen(doc.SelStartLine)
		
        If (doc.SelStartCol = 0) And (doc.SelEndCol = 0) Then ' выделено ровно одна или несколько строк
            Line2 = doc.SelEndLine - 1
            Col2 = doc.LineLen(doc.SelEndLine - 1)
        Else
            Line2 = doc.SelEndLine
            Col2 = doc.LineLen(doc.SelEndLine)
        End If  
		
	    sArray = Split(doc.Range(Line1, Col1, Line2, Col2), vbCrLf)
		CurrentText = ""
	    For Each s In sArray
	        CurrentText = CurrentText & s & vbCrLf
	    Next
	    For Each s In sArray
	        CurrentText = CurrentText & s & vbCrLf
	    Next
	    doc.Range(Line1, Col1, Line2, doc.LineLen(Line2)) = CurrentText
	
	    ' верну курсор в ту же позицию с учетом добавленных строк
	    iNewLine = Line1 + UBound(sArray) + 1
	    doc.MoveCaret iNewLine, doc.SelStartCol, iNewLine, doc.SelStartCol
	
    End If
End Sub ' CopyLine0     

' вспомогательная функция к ExchangeLeftAndRightOfAssign
Function ExchangeLeftAndRightOfAssignInString(Source)
	ExchangeLeftAndRightOfAssignInString = Source

    Set reg = New RegExp
        reg.Pattern = "^(\s*)([^=\s]+)(\s*)=(\s*)([^;]+)(\s*;.*)?"
        reg.IgnoreCase = True
		
    ' если не строка-комментарий
    If CommonScripts.RegExpTest("^\s*//", Source) Then
	  	Exit Function
	end if
    Set Matches = reg.Execute(Source)
    If Matches.Count > 0 Then
		ExchangeLeftAndRightOfAssignInString = Matches(0).SubMatches(0) + Matches(0).SubMatches(4) + Matches(0).SubMatches(2) + "=" + Matches(0).SubMatches(3) + Matches(0).SubMatches(1) + Matches(0).SubMatches(5)
	end if
	
End Function

' обменять местами левую и правую часть присвоения с сохранением возможных комментариев и пробелов возле знака присвоения
' например, вместо
' 		КолСубс1 =  гМассивПараметровДляРасчета[ПРМ_КолСубс1];  // пример 
' будет
' 		гМассивПараметровДляРасчета[ПРМ_КолСубс1] =  КолСубс1;  // пример 

Sub ExchangeLeftAndRightOfAssign()
	
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub 
	
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = ExchangeLeftAndRightOfAssignInString(doc.Range(doc.SelStartLine))
		
  	else ' выделено несколько строк
		if doc.SelEndLine < doc.LineCount-1 then
	    	Source = doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0)
		else
	    	Source = doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))
		end if
	    sArray = Split(Source, vbCrLf)
		Dest = ""
	    For Each s In sArray 
			if s <> "" then
	        	Dest = Dest & ExchangeLeftAndRightOfAssignInString(s) & vbCrLf
			end if
	    Next
		if doc.SelEndLine < doc.LineCount-1 then
			doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) = Dest
		else
			doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine)) = Dest
		end if
  	end if
  	
    doc.MoveCaret doc.SelStartLine, doc.SelStartCol, doc.SelStartLine, doc.SelStartCol
	
End Sub ' ExchangeLeftAndRightOfAssign

' для скриптов VBScript символ комментария это апостроф
Function GetCommentSymbol(doc)
	GetCommentSymbol = "//"
	If (InStrRev(UCase(doc.Path), ".VBS") > 0) then
		GetCommentSymbol = "'"
	end if
End Function

' комментировать текст 
' в отличие от типового ставит комментарий, 
' даже если ничего не выделено или выделен текст на одной строке
' правильно комментирует vbs-скрипты (символом апострофа)
'
Sub CommentSelection()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
		                        
	sCommentSymbol = GetCommentSymbol(doc)

	' если ничего не выделено или выделен текст на одной строке
	if (doc.SelStartLine = doc.SelEndLine) then
    	Set reg = New RegExp
        reg.Pattern = "^(\s*)(.+)"
		Set Matches = reg.Execute(doc.Range(doc.SelStartLine))
		If Matches.Count > 0 Then
			' если строка не пуста, комментарий ставлю после пробельных символов
			doc.Range(doc.SelStartLine) = Matches(0).SubMatches(0) & sCommentSymbol & Matches(0).SubMatches(1)
		else ' если пуста, комментарий ставлю в начале
			doc.Range(doc.SelStartLine) = sCommentSymbol & doc.Range(doc.SelStartLine)
		end if
	    doc.MoveCaret doc.SelStartLine, doc.SelStartCol +Len(sCommentSymbol), doc.SelStartLine, doc.SelEndCol+Len(sCommentSymbol)
	else
  		doc.CommentSel()
  	end if
End Sub
         
' раскомментировать текст 
' в отличие от типового ставит комментарий, 
' даже если ничего не выделено или выделен текст на одной строке
' правильно раскомментирует vbs-скрипты (символом апострофа)
'
Sub UnCommentSelection()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
	sCommentSymbol = GetCommentSymbol(doc)

	' если ничего не выделено или выделен текст на одной строке
	if (doc.SelStartLine = doc.SelEndLine) then
		sText = doc.Range(doc.SelStartLine)
		pos = InStr(sText, sCommentSymbol)
		if pos <> 0 then
			doc.Range(doc.SelStartLine) = Left(sText, pos-1) & Mid(sText, pos+Len(sCommentSymbol))
		end if
	    doc.MoveCaret doc.SelStartLine, doc.SelStartCol-Len(sCommentSymbol), doc.SelStartLine, doc.SelEndCol-Len(sCommentSymbol)
	else
		doc.UnCommentSel()
  	end if
End Sub 

Function GetWordFromString(Source)	
	GetWordFromString = ""
	
    Set reg = New RegExp
        reg.Pattern = "^([^\s;]+)([\s;]*)"
        reg.IgnoreCase = True
		
    Set Matches = reg.Execute(Source)
    If Matches.Count > 0 Then
		GetWordFromString = Matches(0).SubMatches(0)
	end if
End Function     
	                  
' ------------------------------------------------------------------------------------------------
' код для облегчения работы с Телепатом.
'
' Принцип работы:
' когда в конструкции типа 
'      << Перем1 = |Перем2; >>
' где | - положение курсора,
' 
' начинаем набирать первые буквы какой-нибудь функции, например, "Сокр",
' Телепата выводит список возможных вариантов (СокрЛ, СокрЛП, СокрП),
' после выбора, например, СокрЛП, исходный код преобразуется в 
'     << Перем1 = СокрЛП()Перем2; >> // смотрится коряво, и требуется вручную расставлять скобки
'                               
' что выглядит не очень красиво.
' Так вот следующий код позволяет при тех же самых действиях из строки
'     << Перем1 = |Перем2; >>
' получить строку
'     << Перем1 = СокрЛП(Перем2);  >> // выглядит намного лучше, не правда ли
' 
' т.е. то, что в большинстве случае мы и хотим получить 
' 
' PS в следующей версии попытаюсь получить код, который будет позволять отмену
' т.е. после получения << Перем1 = СокрЛП(Перем2); >>, нажав Ctrl-Z (или Действия - Отменить),
' можно будет получить строку << Перем1 = СокрЛП()Перем2); >>
' и таким образом, пользователю легко выбирать нужное ему поведение при вставке функции
' ------------------------------------------------------------------------------------------------
'
Sub InsertMethod(InsertType, InsertName, Text)
	
    Set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
		
	Line = doc.SelStartLine
	Col = doc.SelStartCol
	str1 = doc.Range(Line, 0, Line, Col)
    str2 = doc.Range(Line, Col, Line, doc.LineLen(Line))
	
	str1 = StrReverse(str1)
	 
	str1 = GetWordFromString(str1)
	str1 = StrReverse(str1)
	
	str2 = GetWordFromString(str2)

	if InStr(Text, "!") > 0 then ' СокрЛП(!)
		if str2 <> "" then
			if bDontInsertEndCommaInTelepatReplaceFlag then
				Text = Replace(Text, "!", str2)
			else
				replStr = "!)"
				if InStr(1, Text, "!);") > 0 then
					replStr = "!);"
				end if
				Text = Replace(Text, replStr, "!" & str2)				
			end if
		end if
	else ' КакаяТоФункция()
		if bDontInsertEndCommaInTelepatReplaceFlag then
			Text = Replace(Text, "()", "(" & str2 & ")")
		else
			replStr = "()"
			if InStr(1, Text, "();") > 0 Then
				replStr = "();"
			end if
			Text = Replace(Text, replStr, "(!" & str2)			
		end if
	end if		
	
	doc.Range(Line, Col, Line, Col + Len(str2)) = ""

End Sub

' Обработчик события вставки текста из списка завершения
' Позволяет изменить текст вставки.
' InsertType - тип вставляемого значения (пояснения ниже)
' InsertName - имя вставляемого значения (как оно пишется в списке завершения)
' Text - вставляемый текст
' Во вставляемом тексте местоположение знака "!" определяет размещение
' курсора после вставки. (работает корректно только для однострочных вставок)
' Если положение курсора не указано, то он устанавливается в конце текста.
' При вставке шаблона из списка завершений данный обработчик не вызывается.
' Для примера показано, как вместо И,ИЛИ,НЕ вставлять и,или,не
'
Sub Telepat_OnInsert(InsertType, InsertName, Text)
	
    Select Case InsertType
        Case 2          ' Экспортируемый метод текущего модуля
			InsertMethod InsertType, InsertName, Text
        Case 3          ' Неэкспортируемый метод текущего модуля
			InsertMethod InsertType, InsertName, Text
        Case 7          ' Экспортируемый метод глобального модуля
			InsertMethod InsertType, InsertName, Text
        Case 8          ' Неэкспортируемый метод глобального модуля
			InsertMethod InsertType, InsertName, Text
        Case 10         ' Метод 1С 
			InsertMethod InsertType, InsertName, Text
    End Select
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
End Sub

Init 0
