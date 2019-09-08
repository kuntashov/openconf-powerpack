$NAME ������ ==>> ����� ������
' ������, ������� �������� ����� �������� ������ � ����� ������
'
' �����: ������� ����� aka ADirks
' e-mail: <adirks@sibbereg.ru>
'
'Dim CommonScripts
  
'����������� ������� ������ � ��������
Sub CopyTextToClipboard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub

	' ����������� ������� ������� �������
	Set Pos = CommonScripts.GetDocumentPosition(Doc)
	    
	str=trim(Doc.range(0,0))
	if UCase(left(str,18)) = UCase("#���������������� ") then
		' ����� �������� - ���������
		nNumString = 1
	ElseIf UCase(left(str,14))=UCase("#LoadFromFile ") then
		' � �� �� � ����. ���������
		nNumString = 1
	Else
		nNumString = 0
	End If

	' ������� ���� �����
	Doc.MoveCaret nNumString, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)
	
	' �������� ���������� �������� � ����� ������
	Doc.Copy

	' ��������������� ������ � �������� �������
	Pos.Restore

	Status "����� �������� ������ ���������� � clipboard"
End Sub ' CopyTextToClipboard


' ������, �������  �������� ������� ������ �� ���������� ���������
Sub ReplaceTextFromClipBoard()
	Set Doc = CommonScripts.GetTextDocIfOpened(0)
	If Doc Is Nothing Then Exit Sub
            
	' ����������� ������� ������� �������
	Set Pos = CommonScripts.GetDocumentPosition(Doc)
	
	nNumString = 0
	For I = 0 To Doc.LineCount-1 ' ���������� ������
		str=trim(Doc.range(i,0))
		if UCase(left(str,18)) = UCase("#���������������� ") then
			' ����� �������� - ���������
			nNumString = I+1
			Exit For
		End If
		if UCase(left(str,14))=UCase("#LoadFromFile ") then
			' � �� �� � ����. ���������
			nNumString = I+1
			Exit For
		End If
	Next

	' ������� ���� ����� �������� ������
	Doc.MoveCaret nNumString, 0, Doc.LineCount-1, Doc.LineLen(Doc.LineCount-1)
	
	' ������� ���������� �������� �� ���������� clipboard"
	Doc.Paste
	Pos.Restore

      Status "����� �������� ������ ������ �� ���������� clipboard"
End Sub

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