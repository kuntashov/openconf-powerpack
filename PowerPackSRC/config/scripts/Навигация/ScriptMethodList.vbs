' �������� ������ ������� �������������� ������� (����� *.vb, *.js, *.wsc)
'
'	������ "ShowMethodsList" - � ������ ���������� � ���� ������� (vbs,js,wsc) �������� 
'		������ "ScriptMethodsList" ��� ������ ������� �������, 
'		� ����� ������ ���������� ���� "������ ������" �� ��������.
'	���� ������ ������ �������� �� Ctrl+1 (������ �������� �� ��������)
'
'������: $Revision: 1.6 $ 
'                   
' ������:
'	MetaEditor
'	����� ������� aka artbear
'
'���������� ������, ������� �������� ���������
Const SupportedExtensions = ".vbs.js.wsc"

Sub ShowMethodsList()
	
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit Sub
	
	if CommonScripts.CheckDocOnExtension(doc, SupportedExtensions) Then 
		ScriptMethodsList
	else
		SendCommand(33298) ' �����������������������������������
	end if
End Sub

re_proc = "^\s*(private\s*)?((?:sub)|(?:function))\s+([\w�-���\d]+)\s*\(([\w�-���\d\s,.=""\']*)\)\s*"

'=======================================================================================
Sub ScriptMethodsList()

	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub

	Set MethodsDict = CreateObject("Scripting.Dictionary")

	for i=0 to doc.LineCount-1 
		sLine = lTrim(doc.Range(i))
		set Matches = CommonScripts.RegExpExecute(re_proc, sLine)
		if not Matches is Nothing then 
			
			for each Match in Matches   
				MethodsDict.Add Match.SubMatches(2), i
			next
		end if
	next       
	if MethodsDict.Count = 0 then Exit Sub
		
  	num = CommonScripts.SelectValue(MethodsDict)
  	if num = "" then exit sub

	doc.MoveCaret num,0,num,0
End Sub

'
' ��������� ������������� �������
'
Private Sub Init() ' ��������� ��������, ����� ��������� �� �������� � �������
	
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
	
End Sub ' Init

Init