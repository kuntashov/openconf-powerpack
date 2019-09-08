$NAME Модуль ==>> буфер обмена
' Макрос, который копирует текст текущего модуля в буфер обмена
'
' Автор: Алексей Диркс aka ADirks
' e-mail: <adirks@sibbereg.ru>
'
'Dim CommonScripts
  
'Скопировать текущий модуль в клипборд
Sub CopyTextToClipboard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	' зафиксируем текущую позицию курсора
	Set Pos = CommonScripts.GetDocumentPosition(Doc)
	    
	str=trim(Doc.range(0,0))
	if UCase(left(str,18)) = UCase("#ЗагрузитьИзФайла ") then
		' Нашли загрузку - открываем
		nNumString = 1
	ElseIf UCase(left(str,14))=UCase("#LoadFromFile ") then
		' И то же с англ. вариантом
		nNumString = 1
	Else
		nNumString = 0
	End If

	' выделим весь текст
	Doc.MoveCaret nNumString, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)
	
	' копируем выделенный фрагмент в буфер обмена
	Doc.Copy

	' восстанавливаем курсор в исходной позиции
	Pos.Restore

	Status "Текст текущего модуля скопирован в clipboard"
End Sub ' CopyTextToClipboard


' Макрос, который  заменяет текущий модуль на содержимое клипборда
Sub ReplaceTextFromClipBoard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
            
	' зафиксируем текущую позицию курсора
	Set Pos = CommonScripts.GetDocumentPosition(Doc)
	
	nNumString = 0
	For I = 0 To Doc.LineCount-1 ' Перебираем строки
		str=trim(Doc.range(i,0))
		if UCase(left(str,18)) = UCase("#ЗагрузитьИзФайла ") then
			' Нашли загрузку - открываем
			nNumString = I+1
			Exit For
		End If
		if UCase(left(str,14))=UCase("#LoadFromFile ") then
			' И то же с англ. вариантом
			nNumString = I+1
			Exit For
		End If
	Next

	' выделим весь текст текущего модуля
	Doc.MoveCaret nNumString, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)
	
	' заменим выделенный фрагмент на содержимое clipboard"
	Doc.Paste
	Pos.Restore

      Status "Текст текущего модуля заменён на содержимое clipboard"
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
End Sub
 
Init 0 ' При загрузке скрипта выполняем инициализацию