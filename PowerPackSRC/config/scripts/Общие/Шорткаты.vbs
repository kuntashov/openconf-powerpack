'�������� ���� ShootNICK + ����������� ������ �����, � �.�. � ��� (artbear)
'������� ���� ������������ ������ ���� ���������� �� ��� ��������� ������...

' artbear

'������������ �������. ���� ������ �� ������� - ��������� ���� ������������, ���� ������ ����� - ������ SelectAll, 
'���� ���� �� ��������� - ��������� ��
Sub CtrlA()
	If Windows.ActiveWnd Is Nothing Then
		CommonScripts.SendCommand cmdOpenConfigWnd
		Windows.ActiveWnd.Maximized = true
		OffPanel "���� ���������"
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
	TogglePanel "�������-��������"
End Sub

Sub ToggleOutPutWindow()
	TogglePanel "���� ���������"
End Sub

Sub TogleSearchWindow()
	TogglePanel "������ ��������� ���������"
End Sub

Sub TogleStdToolbar()
	TogglePanel "�����������"
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
  	
  Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1) ' �������� ���
End Sub

'����������� ������� ������ � ��������
Sub CopyTextToClipboard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	Set Pos = CommonScripts.GetDocumentPosition(Doc)

	Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1) ' �������� ���

	Doc.Copy

	Pos.Restore

	Status "����� �������� ������ ���������� � clipboard"
End Sub ' CopyTextToClipboard

'�������� ������� ������ �� ���������� ���������
Sub ReplaceTextFromClipBoard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	Set Pos = CommonScripts.GetDocumentPosition(Doc)

	Doc.MoveCaret 0, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)

	Doc.Paste
	Pos.Restore

	Status "����� �������� ������ ������ �� ���������� clipboard"
End Sub

Sub SyntaxCheck()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
  SendCommand(33297) '�������������� ��������
End Sub    

Sub CloseMessageWindow()                   
	'SendCommand(32812) '������� ���� ���������
	CommonScripts.TogglePanel "���� ���������"
End Sub

'Sub Configurator_AllPluginsInit()
'	CloseMessageWindow
'End Sub
    
' ADirks
Sub Run1CInExlusiveMode() '����������� ������ F11
  SendCommand(33876)
End Sub

Sub OpenInDebugger()
  SendCommand(33285)
End Sub ' OpenInDebugger

' artbear
Sub OpenGlobalModule()
	Documents("����������������").Open
End Sub ' OpenGlobalModule

' ������ ������� ��� �������� ���� �������-��������� �� ������ ���������.
' orefkov
Dim closesp
closesp=false

Sub Configurator_OnIdle()
  if closesp then closesp=false:SendCommand 45098
End Sub

Sub Configurator_OnDoModal(Hwnd, Caption, Answer)
  if closesp and Caption="����������� �����" then Answer=mbaOK:Exit sub
End Sub

Sub SyntaxHelperClose()
  closesp=true
  SendCommand 33879
End Sub ' SyntaxHelperClose 

' a13x, 2005/03/27
' �������� � ������/��������� ������ ��������� ��������� a la Far Manager
' �������� ������������ � ����� ������� ��������!
'
Sub FarCtrlPgUp() ' � ������ ������
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	CommonScripts.Jump 0, Doc.SelStartCol, 0, Doc.SelStartCol
End Sub 

Sub FarCtrlPgDown() ' � ��������� ������
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	LineNo	= Doc.LineCount - 1
	ColNo	= Doc.SelStartCol
	CommonScripts.Jump LineNo, ColNo, LineNo, ColNo
End Sub

' a13x, 2005/03/27
' Ctrl+A � ��� ������������ �������� ("�������� ���")
Sub ClassicCtrlA()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
	LastLine = Doc.LineCount - 1
	Doc.MoveCaret 0, 0, LastLine, Doc.LineLen(LastLine)-1
End Sub

'��������� ������  '<'. ������������ ��� ��������� �� Ctrl-� ��� ������� ��������� ����������
Sub OpenAngleBracket()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = "<"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+1, doc.SelStartLine, doc.SelStartCol+1
End Sub

'��������� ������  '>'. ������������ ��� ��������� �� Ctrl-� ��� ������� ��������� ����������
Sub CloseAngleBracket()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
	doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = ">"
	doc.MoveCaret doc.SelStartLine, doc.SelStartCol+1, doc.SelStartLine, doc.SelStartCol+1
End Sub


'
' ��������� ������������� �������
'
Private Sub Init()
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
 
Init ' ��� �������� ������� ��������� �������������