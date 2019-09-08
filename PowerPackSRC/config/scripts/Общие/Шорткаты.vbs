'исходная идея ShootNICK + добавляения других людей, в т.ч. и мои (artbear)
'Напихал сюда обработчиков модуля чтоб расставить на них сочетания клавиш...

' artbear

'клавиатурный шорткат. Если ничего не открыто - открывает окно конфигурации, если открыт текст - делает SelectAll, 
'если окно не текстовое - открывает ГМ
Sub CtrlA()
	If Windows.ActiveWnd Is Nothing Then
		CommonScripts.SendCommand cmdOpenConfigWnd
		Windows.ActiveWnd.Maximized = true
		OffPanel "Окно сообщений"
		Exit Sub
	End If
	OldMode = CommonScripts.SetQuietMode(true)
	Set Doc = CommonScripts.GetTextDocIfOpened(false, true)
	CommonScripts.SetQuietMode OldMode
	If Doc Is Nothing Then
		OpenGlobalModule
	Else
		'SelectAll
		CopyTextToClipboard
	End If
End Sub


Sub TogglePanel(PanelName)
	Windows.PanelVisible(PanelName)=Not Windows.PanelVisible(PanelName)
End Sub

Sub OffPanel(PanelName)
	Windows.PanelVisible(PanelName) = false
End Sub

Sub OnPanel(PanelName)
	Windows.PanelVisible(PanelName) = true
End Sub

Sub ToggleSyntaxHelper()
	TogglePanel "Синтакс-Помощник"
End Sub

Sub ToggleOutPutWindow()
	TogglePanel "Окно сообщений"
End Sub

Sub TogleSearchWindow()
	TogglePanel "Список найденных вхождений"
End Sub

Sub TogleStdToolbar()
	TogglePanel "Стандартная"
End Sub

Sub CommentSelection()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
		
  doc.CommentSel()
End Sub

Sub UnCommentSelection()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  doc.UnCommentSel()
End Sub 

Sub FormatSelection()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  doc.FormatSel()
End Sub

Sub SelectAll()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1) ' выделить все
End Sub

'Скопировать текущий модуль в клипборд
Sub CopyTextToClipboard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	Set Pos = CommonScripts.GetDocumentPosition(Doc)

	Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1) ' выделить все

	Doc.Copy

	Pos.Restore

	Status "Текст текущего модуля скопирован в clipboard"
End Sub ' CopyTextToClipboard

'Заменить текущий модуль на содержимое клипборда
Sub ReplaceTextFromClipBoard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	Set Pos = CommonScripts.GetDocumentPosition(Doc)

	Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)

	Doc.Paste
	Pos.Restore

	Status "Текст текущего модуля заменён на содержимое clipboard"
End Sub

Sub SyntaxCheck()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  SendCommand(33297) 'синтаксический контроль
End Sub    

Sub CloseMessageWindow()                   
	'SendCommand(32812) 'закрыть окно сообщений
	CommonScripts.TogglePanel "Окно сообщений"
End Sub

'Sub Configurator_AllPluginsInit()
'	CloseMessageWindow
'End Sub
    
' ADirks
Sub Run1CInExlusiveMode() 'стандартная кнопка F11
  SendCommand(33876)
End Sub

Sub OpenInDebugger()
  SendCommand(33285)
End Sub ' OpenInDebugger

' artbear
Sub OpenGlobalModule()
	Documents("ГлобальныйМодуль").Open
End Sub ' OpenGlobalModule

' Пример скрипта для закрытия окна синтакс-помощника из любого положения.
' orefkov
Dim closesp
closesp=false

Sub Configurator_OnIdle()
  if closesp then closesp=false:SendCommand 45098
End Sub

Sub Configurator_OnDoModal(Hwnd, Caption, Answer)
  if closesp and Caption="Контекстный поиск" then Answer=mbaOK:Exit sub
End Sub

Sub SyntaxHelperClose()
  closesp=true
  SendCommand 33879
End Sub ' SyntaxHelperClose 

' a13x, 2005/03/27
' Переходы к первой/последней строке открытого редактора a la Far Manager
' Переходы запоминаются в стеке прыжков Телепата!
'
Sub FarCtrlPgUp() ' к первой строке
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	CommonScripts.Jump 0, Doc.SelStartCol, 0, Doc.SelStartCol
End Sub 

Sub FarCtrlPgDown() ' к последней строке
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	LineNo	= Doc.LineCount - 1
	ColNo	= Doc.SelStartCol
	CommonScripts.Jump LineNo, ColNo, LineNo, ColNo
End Sub

' a13x, 2005/03/27
' Ctrl+A в его традиционном варианте ("Выделить все")
Sub ClassicCtrlA()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	LastLine = Doc.LineCount - 1
	Doc.MoveCaret 0, 0, LastLine, Doc.LineLen(LastLine)-1
End Sub

'Вставляет символ  '<'. Предназначен для повешения на Ctrl-Б при русской раскладке клавиатуры
Sub OpenAngleBracket()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = "<"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+1, doc.SelStartLine, doc.SelStartCol+1
End Sub

'Вставляет символ  '>'. Предназначен для повешения на Ctrl-Ю при русской раскладке клавиатуры
Sub CloseAngleBracket()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = ">"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+1, doc.SelStartLine, doc.SelStartCol+1
End Sub


'
' Процедура инициализации скрипта
'
Private Sub Init()
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
 
Init ' При загрузке скрипта выполняем инициализацию