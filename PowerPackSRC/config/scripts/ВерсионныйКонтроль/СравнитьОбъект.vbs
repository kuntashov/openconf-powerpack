$NAME �������� ������ � ���������� �������

'    ��������� �������� ������������ ��� ������ gcomp
'
' ���������� ������� ������ (����������, ��������, �����) � ���, ��� ����� � �������� Src.
' ��� ��������� ����������� ���������� ������������ �������� ������� � ������� Src_New,
' � ����� ����������� <DiffCommand> Src Src_New
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
	
	'����� �������    ����������.������������.�����
	RE.Pattern = "([^\s.]+)\.([^\s.]+)\..*"
	match = RE.Test(Doc.Name)
	If not match Then
		'������ �������������� ������� �������    CMDSubDoc::���������� ������������
		RE.Pattern = "CMD[a-zA-Z]+Doc::(\S+)\s+(\S+)"
		match = RE.Test(Doc.Name)
	End If
	If not match Then
		'������������ CMDTabDoc::������������
		RE.Pattern = "CMD[a-zA-Z]+Doc::(\S+)(.*)"
		match = RE.Test(Doc.Name)
		If match Then
			'Message "MDWnd = " & MDWnd.GetSelected
			parts = split(MDWnd.GetSelected, "\")
			'Message "UBound = " & UBound(parts)
			If parts(0) = "" Then Exit Function
			GetObjPath = parts(0)
			If parts(0) = "����� ������" Then
				GetObjPath = "��������������������"
			ElseIf parts(0) = "���� ��������" Then
				GetObjPath = "��������������������"
			ElseIf parts(0) = "��������" Then
				GetObjPath = "��������������������"
			ElseIf parts(0) = "��������" Then
				GetObjPath = "��������������������"
			ElseIf parts(0) = "���������" Then
				GetObjPath = parts(0)
			ElseIf parts(0) = "������������" Then
				GetObjPath = parts(0)
			ElseIf parts(0) = "��������" Then
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
		Case "����������" aType = "�����������"
		Case "��������"   aType = "���������"
		Case "�������"    aType = "��������"
			aKind = ""
		Case "�����"      aType = "������"
		Case "���������"  aType = "���������"
		Case "������������"  aType = ""
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

