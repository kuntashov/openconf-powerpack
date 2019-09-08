'$NAME ��������� �� ������
' ������ ������� � ������/���������� ������� ���������/�������
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'

Dim PositionInModule

Sub GotoBeginOfMethod()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	set PositionInModule =  CommonScripts.GetDocumentPosition(doc)

	ModuleText = split(doc.Text, vbCrLf)
	for i = doc.SelStartLine to 0 step -1
		sText = UCase(ModuleText(i))
		if Instr(sText,"���������") = 1 or Instr(sText,"�������") = 1 then
			doc.MoveCaret i, 0
			Exit For
		end if
	next
End Sub ' GotoBeginOfMethod

Sub GotoEndOfMethod()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	set PositionInModule =  CommonScripts.GetDocumentPosition(doc)

	ModuleText = split(doc.Text, vbCrLf)
	for i = doc.SelStartLine to UBound(ModuleText)
		sText = UCase(ModuleText(i))
		if Instr(sText,"��������������") = 1 or Instr(sText,"������������") = 1 then
			doc.MoveCaret i, 0
			Exit For
		end if
	next
End Sub ' GotoEndOfMethod

Sub ReturnToSavedPosition()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	PositionInModule.Restore

End Sub ' ReturnToSavedPosition 

' ADirks
Sub SelectProcedure()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  
  GotoBeginOfMethod()
  l1 = Doc.SelStartLine
  GotoEndOfMethod()
  l2 = Doc.SelStartLine
  
  Doc.MoveCaret l1, 0, l2+1, 0
End Sub ' SelectProcedure

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
End Sub
 
Init 0 ' ��� �������� ������� ��������� �������������