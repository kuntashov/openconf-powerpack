'Intellisence - ������ ����� �����
'
'������: $Revision: 1.55 $
'
'����� ������ �������, 2004-2005 ���
'	dimoff66@mail.ru
'��������:
'	������� ����� aka ADirks
'	����� ������� aka artbear
'	��������� �������� aka a13x
'	������� ������ aka trdm
'
'TODO ���� ��� ���������� ����:
'1) �������� �������� ���-����� ��� ������� 1�++,
'� �������� trdm ��� ������ ����� � ������� ��� ��������� ��������� �����,
'������ ���� �������� �� ����� :) -
'����� ������ ����� ������ � ���������� ������ � als-������.
'
'2) ��� ����� ��������� �������� �� ADirks �� ������� ������� � ������� 1C++ �� ������� ��������� ��������� � ������� � ��������� ����� ���������� ������.
'
'3)� ��� ����� �������� ���/���-���� ��� ������ � ����� �������� ���-���� ��� �������� :)
'
'4) � ��� ��� �� ������ ��������� ���/���-������ ��� �������� �� ���� ����� ���-������, �? ;-)
'
'5) ���������� ��������� Intellisense.vbs ����������, � ���������� ��� ������� �������������� ��� ����������. ������ �� �������� ����������� ������������ ����� ������� :)
'TODO end

'���������� ������, ��� ������� �� ����� ��������� �����������
Const NotIntellisenceExtensions = ".vbs.js.wsc.prm"

' ��������� ��� ������ ��� ������ ���������� ���������
' ��� ��� ������ ����� � �������
Dim DebugFlag '����������� ���������� ����������
DebugFlag = True '�������� ����� ���������� ���������
DebugFlag = false '�������� ����� ���������� ���������

'����������� ����������
Const picAttr           = "3 #"
Const picMeth           = "7 #"
Const picProcs          = "1 #"
Const picObjField       = "4 #"
Const picObjTableField  = "5 #"
Const picFormAttr       = "6 #"
Const picTableField     = "10#"
Const picCommonDocField = "4 #"
Const PicMetaDataKind   = "8 #"


Dim DataDir

Dim CreateObjectFlag
Dim doc

Dim txtStrings
Dim txtTrim
Dim txtLow
Dim ModuleWnd_HWnd, ActiveWnd_HWnd
Dim CurrentLine
Dim CurDocLine, CurDocCol
Dim ToAddDot

Dim IdentLine
Dim SymbBefore
Dim SymbAfter
Dim FinalObjType
Dim TypeDef_Pos '������� ���������, ������������� ��� (� txtLow)

Dim strPossibleTypes
Dim MainTypes
Dim MethodsBehaviour
Dim TypesStructure
Dim TypesStructureExt
Dim MethodsParamType
Dim IniParams
Dim OtherTypesDefine
Dim CPPAnalyser

Dim DlgType

Dim FSO

Dim gTimerIns		' ������ �������� ����� ������� ������ OnInsert

Set MainTypes         = CreateObject("Scripting.Dictionary")
Set MethodsBehaviour  = CreateObject("Scripting.Dictionary")
Set TypesStructure    = CreateObject("Scripting.Dictionary")
Set TypesStructureExt = CreateObject("Scripting.Dictionary")
Set MethodsParamType  = CreateObject("Scripting.Dictionary")
Set IniParams         = CreateObject("Scripting.Dictionary")
Set OtherTypesDefine  = CreateObject("Scripting.Dictionary")
Set KeyMeths          = CreateObject("Scripting.Dictionary")

Set CPPAnalyser = Nothing
On Error Resume Next
Set CPPAnalyser = CreateObject("OpenConf._1CPP")
If Not CPPAnalyser Is Nothing Then CPPAnalyser.SetConfig(Configurator)
On Error Goto 0

' �������� ����������� � ������ ��� �����
' ������ ����� ����������� �� �������� MethodsBehaviour
'
MainTypes.Add "������������", "������������������"
MainTypes.Add "��������", "��������������"
MainTypes.Add "����������", "����������������"
MainTypes.Add "������������", "������������������"
MainTypes.Add "�������", "�������������"
MainTypes.Add "���������", "���������������"
MainTypes.Add "����������", "�����������������"
MainTypes.Add "�����", "�����#1"
MainTypes.Add "��������", "��������#2"
MainTypes.Add "��������������", "FRMTYPE#3"
MainTypes.Add "���������������", "FRMTYPE#4"
MainTypes.Add "��", "��"
MainTypes.Add "����������", "����������"
MainTypes.Add "����������", "����"
MainTypes.Add "��������������", "������������"
MainTypes.Add "��������������", "�������������������"

TypesStructure.Add "��������", "ATTR|2|1"
TypesStructure.Add "����������", "ATTR|1|0"
TypesStructure.Add "�������", "ATTR|7|2"
TypesStructure.Add "������������������", "KINDNEW|���������������"
TypesStructure.Add "����������������", "KINDNEW|�������������"
TypesStructure.Add "��������������", "KINDNEW|�����������"
TypesStructure.Add "���������������", "KINDNEW|���������"
TypesStructure.Add "�������������", "KINDNEW|�������"
TypesStructure.Add "���������������", "KINDOLD|������������"
TypesStructure.Add "�����", "VALUE|������������"
TypesStructure.Add "���������������", "VTCOLUMN"
TypesStructure.Add "����������������������", "VTCOLUMN"
TypesStructure.Add "������", "QUERYATTR"
TypesStructure.Add "������������������", "VALUE|�����������"
TypesStructure.Add "���������������", "SUBKREF"
TypesStructure.Add "xbase", "SUBFILD"



TypesStructureExt.Add "��������.���������������", "KINDOLD|��������"
TypesStructureExt.Add "��������.��������", "VALUE|��������"
TypesStructureExt.Add "��������.���", "VALUE|������������������"
TypesStructureExt.Add "����������.��������������", "KINDOLD|����������"
TypesStructureExt.Add "����������.��������", "KINDOLD|����������"
TypesStructureExt.Add "����������.��������", "REFOWNER"
TypesStructureExt.Add "����������.���", "VALUE|��������������������"
TypesStructureExt.Add "������������������.����", "VALUE|����"
TypesStructureExt.Add "������������������.������", "VALUE|����������.������"
TypesStructureExt.Add "������������������.��������", "VALUE|��������"
TypesStructureExt.Add "�����.��������", "VALUE|��������������"
TypesStructureExt.Add "�����.��������", "VALUE|"
TypesStructureExt.Add "�����.���������������", "VALUE|�������������������"
TypesStructureExt.Add "��������.�����", "VALUE|���������������"
TypesStructureExt.Add "��������.������", "VALUE|���������������"
TypesStructureExt.Add "��������.��������", "VALUE|��������"
TypesStructureExt.Add "����.��������", "VALUE|����"
TypesStructureExt.Add "���������������.����", "VALUE|����"
TypesStructureExt.Add "������������.�������������", "KINDOLD|��������������������"
TypesStructureExt.Add "��������������.��������", "VALUE|��������"
TypesStructureExt.Add "��������������.��������������������", "VALUE|��������"
TypesStructureExt.Add "��������������.�������", "VALUE|����������"
TypesStructureExt.Add "����������.��������", "VALUE|������������������"
TypesStructureExt.Add "����������.����������", "VALUE|��������������������"
TypesStructureExt.Add "����������.�������", "VALUE|�����������������"
TypesStructureExt.Add "����������.����������", "VALUE|��������������������"
TypesStructureExt.Add "����������.�����������", "VALUE|���������������������"
TypesStructureExt.Add "����������.�����������", "VALUE|���������������������"
TypesStructureExt.Add "����������.��������������", "VALUE|������������������������"
TypesStructureExt.Add "����������.������", "VALUE|����������������"
TypesStructureExt.Add "����������.��������������", "VALUE|������������������������"
TypesStructureExt.Add "����������.���������", "VALUE|�������������������"
TypesStructureExt.Add "����������.���������", "VALUE|�������������������"
TypesStructureExt.Add "����������.���������", "VALUE|�������������������"
TypesStructureExt.Add "����������.������������", "VALUE|����������������������"
TypesStructureExt.Add "����������.����������", "VALUE|��������������������"
TypesStructureExt.Add "����������.������������������", "VALUE|����������������������������"
TypesStructureExt.Add "��������������������.��������", "VALUE|����������������������������"
TypesStructureExt.Add "�����������������.���������", "VALUE|��������������������������"
TypesStructureExt.Add "�����������������.������", "VALUE|�����������������������"
TypesStructureExt.Add "�����������������.��������", "VALUE|�������������������������"
TypesStructureExt.Add "���������������������.������", "VALUE|���������������������������"
TypesStructureExt.Add "��������������������.�������������������", "VALUE|���������������������������������������"
TypesStructureExt.Add "��������������������.�����������������������", "VALUE|�������������������������������������������"
TypesStructureExt.Add "��������������������.��������������", "VALUE|����������������������������������"
TypesStructureExt.Add "���������������������.������", "VALUE|���������������������������"
TypesStructureExt.Add "������������������������.����������", "VALUE|����������������������������������"
TypesStructureExt.Add "������������������.����������������������������", "VALUE|����������������������������������������������"
TypesStructureExt.Add "������������������������.��������", "VALUE|��������������������������������"
TypesStructureExt.Add "����������������������������.����������������", "VALUE|��������������������������������������������"
TypesStructureExt.Add "����������������������������.���������", "VALUE|�������������������������������������"
TypesStructureExt.Add "����������������������������.�����������������", "VALUE|���������������������������������������������"
TypesStructureExt.Add "����������������������������.���������������������", "VALUE|�������������������������������������������������"
TypesStructureExt.Add "������������������.����������������������", "VALUE|����������������������������������������"
TypesStructureExt.Add "������������������.�������������", "VALUE|�������������������������������"
TypesStructureExt.Add "����������������.�����", "VALUE|���������������������"
TypesStructureExt.Add "����������������.�����������", "VALUE|���������������������������"
TypesStructureExt.Add "������������������������.��������", "VALUE|��������������������������������"
TypesStructureExt.Add "������������������������.�����������", "VALUE|�����������������������������������"
TypesStructureExt.Add "����������������������.��������", "VALUE|������������������������������"
TypesStructureExt.Add "��������������������.����", "VALUE|������������������������"
TypesStructureExt.Add "��������������������.�����������", "VALUE|�������������������������������"
TypesStructureExt.Add "����������.������������������", "VALUE|����������������������������"
TypesStructureExt.Add "����������.���������", "VALUE|�������������������"
TypesStructureExt.Add "����������.����������������������", "VALUE|��������������������������������"
TypesStructureExt.Add "����������.�����", "VALUE|���������������"
TypesStructureExt.Add "����������.������������������", "VALUE|����������������������������"
TypesStructureExt.Add "����������.����������������", "VALUE|��������������������������"
TypesStructureExt.Add "����������.�������������", "VALUE|�����������������������"
TypesStructureExt.Add "����������.���������������", "VALUE|�������������������������"
TypesStructureExt.Add "����������.��������������������", "VALUE|������������������������������"
TypesStructureExt.Add "����������.��������������������", "VALUE|������������������������������"
TypesStructureExt.Add "����������.�����������������", "VALUE|���������������������������"

' ��� ��������� ������ �����/��������� �����������
' ATTR|2|1 = �������� ����������
' KINDS|9 = ��� ���� ��������
' ������ �������� (��������, 2, 1) ������������� � MetaData.TaskDef.Childs(ObjNumber)
'
MethodsBehaviour.Add "��������", "ATTR|2|1"
MethodsBehaviour.Add "����������", "ATTR|1|0"
MethodsBehaviour.Add "�����������", "ATTR|2|1"
MethodsBehaviour.Add "�������������", "ATTR|1|0"
MethodsBehaviour.Add "�������", "ATTR|7|2"
MethodsBehaviour.Add "������������������", "KINDS|4"
MethodsBehaviour.Add "���������������", "KINDS|0"
MethodsBehaviour.Add "����������������", "KINDS|1"
MethodsBehaviour.Add "��������������", "KINDS|2"
MethodsBehaviour.Add "�������������", "KINDS|7"
MethodsBehaviour.Add "��������", "KINDS|24"
MethodsBehaviour.Add "�����������������", "KINDS|9"
MethodsBehaviour.Add "������������������", "KINDS|18"
MethodsBehaviour.Add "���������������", "KINDS|18"
MethodsBehaviour.Add "���������������", "ATTR|4|0"
MethodsBehaviour.Add "���������������", "VTCOLUMNS"
MethodsBehaviour.Add "����������������������", "VTCOLUMNS"
MethodsBehaviour.Add "�����", "DLGCONTROLS"
MethodsBehaviour.Add "������", "QUERYATTRS"
MethodsBehaviour.Add "�������������������", "KINDS|10"
MethodsBehaviour.Add "xbase", "VTFIELDS"
MethodsBehaviour.Add "��������������", "ATTR|11|0"
MethodsBehaviour.Add "������������", "������������|0"

MethodsParamType.Add "��������.����.1", "DOCTABLE"
MethodsParamType.Add "���������������.����.1", "VTCOLUMNS"
MethodsParamType.Add "���������������.�������������.3", "VTCOLUMNS"
MethodsParamType.Add "���������������.����������������.2", "VTCOLUMNS"
MethodsParamType.Add "���������������.������������������.2", "VTCOLUMNS"
MethodsParamType.Add "���������������.��������������������������.1", "VTCOLUMNS"
MethodsParamType.Add "���������������.������������������������.1", "VTCOLUMNS"
MethodsParamType.Add "���������������.�����������.1", "VTCOLUMNS"
MethodsParamType.Add "���������������.��������������.1", "VTCOLUMNS"
MethodsParamType.Add "���������������.��������.1", "VTCOLUMNSCOLL"
MethodsParamType.Add "���������������.��������.2", "VTCOLUMNSCOLL"
MethodsParamType.Add "���������������.����������������.1", "VTCOLUMNSCOLL"

MethodsParamType.Add "����������������������.����.1", "VTCOLUMNS"
MethodsParamType.Add "����������������������.�����������.3", "VTCOLUMNS"
MethodsParamType.Add "����������������������.����������������.2", "VTCOLUMNS"
MethodsParamType.Add "����������������������.������������������.2", "VTCOLUMNS"
MethodsParamType.Add "����������������������.�����������.1", "VTCOLUMNS"
MethodsParamType.Add "����������������������.��������������.1", "VTCOLUMNS"
MethodsParamType.Add "����������������������.��������.1", "VTCOLUMNSCOLL"
MethodsParamType.Add "����������������������.��������.2", "VTCOLUMNSCOLL"

MethodsParamType.Add "����������.���������.1", "KINDS|0"
MethodsParamType.Add "����������.����������.1", "KINDS|1"
MethodsParamType.Add "����������.��������.1", "KINDS|2"
MethodsParamType.Add "����������.�������.1", "KINDS|7"

OtherTypesDefine.Add "v77.application", "v77_application"
OtherTypesDefine.Add "excel.application", "excel_application"
OtherTypesDefine.Add "word.application", "word_application"


DataDir = BinDir & "Config\Intell"
strPossibleTypes = ""

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'������ ���-�����
'
Set FSO = CreateObject("Scripting.FileSystemObject")
IniFileName = DataDir + "\intell.ini"
if FSO.FileExists(IniFileName) then
	Set Fl = FSO.GetFile(IniFileName)
	Set FileStream = Fl.OpenAsTextStream()
	AllPrm = FileStream.ReadAll()
	AllParams = Split(AllPrm, vbCrLf)
	for i = 0 to UBound(AllParams)
		if left(AllParams(i), 2) <> "//" and instr(AllParams(i), "=") > 0 then
			IniParams.Add Trim(leftFrom(AllParams(i), "=")), Trim(RightFrom(AllParams(i), "="))
		end if
	next

	if UCase(GetIniParam("TELEPAT")) = "DEFAULT" then IniParams.Add "TELEPAT", "��"
	if UCase(GetIniParam("SHOWTOOLTIPS")) = "DEFAULT" then IniParams.Add "SHOWTOOLTIPS", "��"
end if

'������ ������ ���������� ����� ��������
UniqueFileName = DataDir + "\Key Methods.ints"
if FSO.FileExists(UniqueFileName) then
	Set Fl = FSO.GetFile(UniqueFileName)
	Set FileStream = Fl.OpenAsTextStream()
	AllPrm = FileStream.ReadAll()
	AllParams = Split(AllPrm, vbCrLf)
	for i = 0 to UBound(AllParams)
		KeyCouple = Trim(AllParams(i))
		Pos = instr(KeyCouple, " ")
		if Pos > 0 then
			KeyMeths.Add left(KeyCouple, Pos - 1), Mid(KeyCouple, Pos + 1)
		end if
	next
end if

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� True ���� ������� ���� - ���������
'
Function CheckWindow(doc)
	CheckWindow = False

	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is nothing then exit Function

    if CommonScripts.CheckDocOnExtension(doc, NotIntellisenceExtensions) Then Exit Function

	CheckWindow = True
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���������� �������� ��������� ParamName �� ����� �������������
' ��� ������ "default" ���� �������� �� ������
'
Function GetIniParam(ParamName)
	GetIniParam = "default"

	if IniParams.Exists(ParamName) then
		GetIniParam = IniParams.Item(ParamName)
		if lCase(GetIniParam) = "default" then
			GetIniParam = "default"
		end if
	end if
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ��������� �����, ��������������� ������� Expressions
'� ����� TxtBlock
'
Private Function GetREQuery(TxtBlock, Expression)
	Set Query = New RegExp
	Query.IgnoreCase = True
	Query.Global = True

	On Error Resume Next
	Query.Pattern = Expression
	if Err.Number <> 0 then
		On Error goto 0
		Debug "GetREQuery Expression", Expression
		Debug "GetREQuery TxtBlock", Left(TxtBlock, 50)
		Debug "Err", Err.Description
	end if
	On Error goto 0

	On Error Resume Next
	Set GetREQuery = Query.Execute(TxtBlock)
	if Err.Number <> 0 then
		On Error goto 0
		Debug "GetREQuery.Execute Expression", Expression
		Debug "GetREQuery.Execute TxtBlock", Left(TxtBlock, 100)
		Debug "Err", Err.Description
	end if
	On Error goto 0
End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
'
'	�������������� ����� ���� ������ � ������������ � ����������� "DLLKIND"
'	���������� ����� ������������
'
Private Function SelectFromList(strList, InitialText)

	ExtParamsVal = ""
	FontParam = GetIniParam("FONT")
	if FontParam <> "default" then
		ExtParamsVal = ExtParamsVal + ";Font=" + FontParam
	end if
	FontSizeParam = GetIniParam("FONTSIZE")
	if FontSizeParam <> "default" then
		ExtParamsVal = ExtParamsVal + ";FontSize=" + FontSizeParam
	end if
	if ExtParamsVal <> "" then
		ExtParamsVal = Mid(ExtParamsVal, 2)
	end if

	DLLKIND = GetIniParam("CHOICELIST_KIND")
	Select Case DLLKIND

	Case "2": ' "Svcsvc.Service"
		Arr = Split(strList, vbCrLf)
		for i = 0 to UBound(Arr)
			Arr(i) = Mid(Arr(i), 4)
		Next
		strList = Join(Arr, vbCrLf)

		Set SelObjSvc = CreateObject("Svcsvc.Service")
		SelectFromList = SelObjSvc.FilterValue(strList, 1 + 4 + 16)

	Case "3": ' "SelectDialog.Selection"
		Set SelObj = CreateObject("SelectDialog.Selection")
		SelObj.ListToChoose = strList
		On Error Resume Next
		SelObj.InitialText = "" & InitialText
		On Error Goto 0
		ModuleWnd_HWnd = Windows.ActiveWnd.HWnd

		t = SelObj.WndDescrsSet(0, ModuleWnd_HWnd)
		SelectFromList = SelObj.DoSelection

	Case else: ' "SelectValue.SelectVal"
		Set SelObj = CreateObject("SelectValue.SelectVal")
		ModuleWnd_HWnd = Windows.ActiveWnd.HWnd

		SelectFromList = SelObj.SelectPopUpImg(strList, ModuleWnd_HWnd, vbCrLf, ExtParamsVal)

	End Select

End Function ' SelectFromList

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ��� ������� ����������, � �������� ��������� ����
'
Private Function GetWindType()
	WndNameParts = split(doc.Name, ".")
	if UBound(WndNameParts) > 0 then
		GetWindType = lCase(WndNameParts(0)) + "." + lCase(WndNameParts(1))
	else
		GetWindType = ""
	end if
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ����� ����� ������ s1 �� ����� ��������� � �� s2
'
Private Function LeftFrom(s1, s2)
	If InStr(s1, s2) > 0 Then
		LeftFrom = Left(s1, InStr(s1, s2) - 1)
	Else
		LeftFrom = s1
	End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ����� ������ s1 �� ����� ��������� � �� s2
'
Private Function RightFrom(s1, s2)
	If InStr(s1, s2) > 0 Then
		RightFrom = Mid(s1, InStr(s1, s2) + 1)
	Else
		RightFrom = s1
	End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function strBetween(TheStr, FirstSymb, LastSymb)
	TheString = Mid(TheStr, InStr(TheStr, FirstSymb) + Len(FirstSymb))
	strBetween = LeftFrom(TheString, LastSymb)
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Function RemoveComments(txtStrings)
	RemoveComments = ""
	MaxI = UBound(TxtStrings) - 1
	For i = 0 To MaxI Step 1
		str = TxtStrings(i)
		pos = InStr(str, "//")
		If pos > 0 Then TxtStrings(i) = Left(str, pos-1)
	Next
End Function

Private Function PrepareModuleText(Line, Col)
ProfilerEnter("PrepareModuleText 1")
	txtStrings = split(doc.text, vbCrLf)
ProfilerLeave("PrepareModuleText 1")
ProfilerEnter("PrepareModuleText 2")
	RemoveComments(txtStrings)
ProfilerLeave("PrepareModuleText 2")

ProfilerEnter("PrepareModuleText 3")
	txtTrim = Join(txtStrings, vbCrLf)
ProfilerLeave("PrepareModuleText 4")
ProfilerEnter("PrepareModuleText 5")
	txtTrim = Replace(txtTrim, " ", "")
ProfilerLeave("PrepareModuleText 5")
ProfilerEnter("PrepareModuleText 6")
	txtTrim = Replace(txtTrim, vbTab, "")
ProfilerLeave("PrepareModuleText 6")
ProfilerEnter("PrepareModuleText 7")
	txtLow = LCase(txtTrim)
ProfilerLeave("PrepareModuleText 8")

ProfilerEnter("PrepareModuleText end")
	if UBound(TxtStrings) < 0 then
		CurrentLine = ""
	else
		CurrentLine = Trim( left(txtStrings(Line), Col) )
	end if

	' kuntashov@yandex.ru, 2005/06/04
	Do While (Right(CurrentLine, 1)<>".") and (Len(CurrentLine)>0)
		RSide = Right(CurrentLine,1) + RSide
		' TODO ���� ��� �� ��������� ����� ��������������, �� ����� �� �����
		CurrentLine = Left(CurrentLine, Len(CurrentLine)-1)
	Loop

	If CurrentLine = "" Then
		CurrentLine = RSide
		RSide = ""
	end if
	'/kuntashov@yandex.ru, 2005/06/04

	ToAddDot = true
	Do While Right(CurrentLine, 1) = "."
		CurrentLine = Left(CurrentLine, Len(CurrentLine)-1)
		ToAddDot = false
	Loop

	SymbBefore = ""
	If ToAddDot Then SymbBefore = "."
	SymbAfter = ""
	PrepareModuleText = RSide
ProfilerLeave("PrepareModuleText end")
End Function

'���������� ������ ������� 1�++. ����������� - ','
Private Function CppClassList()
	CppClassList = ""
	If CPPAnalyser Is Nothing Then Exit Function

	If CPPAnalyser.ClassCount = 0 Then CPPAnalyser.UpdateClassesInfo()
	Comma = ""
	Dim i
	For i = 0 to CPPAnalyser.ClassCount - 1
		CppClassList = CppClassList & Comma & CPPAnalyser.ClassInfo(i).Name
		Comma = ","
	Next
End Function

Function LineNumber(pos)
	Dim LeftPart, a
	LeftPart = Left(txtLow, pos + 1)
	a = Split(LeftPart, vbCrLf)
	LineNumber = Ubound(a)
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'
'	������� ������� ������, � ������� ����� ������� ������ ���� ��������� ����� + ������ 1�++
'
Private Function PossibleTypes()
	PossibleTypes = strPossibleTypes
	If strPossibleTypes <> "" Then Exit Function

	CreateObjectFlag = "1"
	If not ReadMethodsFromFile(strPossibleTypes, "CreateObject", "") Then
		ReadMethodsFromFile strPossibleTypes, "�������������", ""
	End If
	strPossibleTypes = Replace(strPossibleTypes, picAttr, "")
	strPossibleTypes = Replace(strPossibleTypes, vbCrLf, ",")

	If strPossibleTypes = "" Then strPossibleTypes = "����������,��������,������������,������,�����,����,�������,���������,����������,��������������,����������,����,��������,������������������,�������,�����,������,��������������,���������������,�������������,��������,�����������������,OLE"

	strPossibleTypes = strPossibleTypes & "," & CppClassList() ' ����� ����� ������� ������ 1�++

	PossibleTypes = strPossibleTypes
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
'
'	������� ������� ������, � ������� ����� vbCrLF ������ ���� ��������� ����� + ������ 1�++
'
Private Function GetAllPossibleTypesList()
	GetAllPossibleTypesList = picMetaDataKind + replace(PossibleTypes(), ",", vbCrLf + picMetaDataKind)
End Function ' GetPossibleTypesList

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
'
'	�������� ��� ������ ������ ���� ��������� ����� + ������ 1�++
'	���������� ����� ������������
'
Private Function SelectTypeFromAllPossibleTypesList()
	Dim strAllPossibleTypes

	strAllPossibleTypes = GetAllPossibleTypesList()

	SelectTypeFromAllPossibleTypesList = SelectFromList(strAllPossibleTypes, "")

End Function ' SelectTypeFromAllPossibleTypesList

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ��������� ������� QueryName, ���� AttributeName - ������ ������
'��� ��� ��������, ���� � AttributeName �������� ��� ���������� �������
'
Private Function GetQueryAttrs(QueryName, AttributeName)
	GetQueryAttrs = ""

	DoQueryStr = QueryName + ".���������("
	Pos = Instr(txtLow, DoQueryStr)
	if Pos = 0 then exit function

	Pos2 = Instr(Pos + Len(DoQueryStr) - 1, txtLow, ")")
	If Pos2 = 0 Then exit function

	QueryText = Mid(txtTrim, Pos + Len(DoQueryStr), Pos2 - Pos - Len(DoQueryStr))

	if Left(QueryText, 1) <> """" then
		QueryVar = QueryText + "\s*=(\s*" + QueryText + "\s*\+\s*)?(((""[^""]+"")|([^"";]))*);"
		Set Matches = GetREQuery(txtTrim, QueryVar)
		QueryText = ""
		for each NextMatch in Matches
'			If NextMatch.Submatches(0) = "" Then
'				'Message "-------------------------------------------"
'				QueryText = ""
'			End If
			Set Matches1 = GetREQuery(NextMatch.Submatches(1), """([^""]+)""")
			For Each Match1 in Matches1
				'Message Match1.Submatches(0)
				QueryText = QueryText + ";" + Match1.Submatches(0)
			Next
		Next

		if QueryText = "" then
			Exit Function
		end if
	end if

	QueryText = Replace(QueryText, vbCrLf, ";")
	QueryText = Replace(QueryText, "|", ";")

	QueryLines = Split(QueryText, ";")
	for i = 0 to UBound(QueryLines)
		Pos = Instr(QueryLines(i), "=")
		if Pos > 0 then
			AttrName = leftFrom(QueryLines(i), "=")
			if Instr(AttrName, "(") = 0 then
				if AttributeName = "" then
					if Instr(lCase(AttrName), "�������") = 1 then
						AttrName = Mid(AttrName, 8)
					end if
					AttrName = PicAttr + AttrName
					if Instr(GetQueryAttrs + vbCrLf, vbCrLf + AttrName + vbCrLf) = 0 then
						GetQueryAttrs = GetQueryAttrs + vbCrLf + AttrName
					end if

				elseif lCase(AttrName) = AttributeName then
					TypeDefine = RightFrom(lCase(QueryLines(i)), "=")
					DefineParts = Split(TypeDefine, ".")
					if UBound(DefineParts) < 2 then
						Exit function
					end if
					DefineParts(2) = LeftFrom(DefineParts(2), ",")

					DefInfo = "_"
					if TypesStructureExt.Exists(DefineParts(0) + "." + DefineParts(2)) then
						DefInfo = TypesStructureExt.Item(DefineParts(0) + "." + DefineParts(2))
					elseif TypesStructure.Exists(DefineParts(0)) then
						DefInfo = TypesStructure.Item(DefineParts(0))
					end if

					DefInfoParts = Split(DefInfo, "|")

					Select case DefInfoParts(0)
					case "KINDOLD":
						GetQueryAttrs = DefInfoParts(1) + "." + DefineParts(1)
					case "ATTR":
						GetQueryAttrs = GetAttrType(CLng(DefInfoParts(1)), DefineParts(1), DefineParts(2), CLng(DefInfoParts(2)))
					case "REFOWNER"
						GetQueryAttrs = GetRefOwnerType(DefineParts(1))
					end select

					exit function
				end if
			End if
		end if
	Next

	if GetQueryAttrs <> "" then GetQueryAttrs = Mid(GetQueryAttrs, 3)

End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ��� ������� �� ����� ������� �������� VTName �
'����� ������� NameToSearch.
'
Private function GetVTColumnType(VTName, NameToSearch)

	GetVTColumnType = ""
	Pos = 1
	do
		NewColumnString = VTName + ".������������("
		Pos = instr(Pos, txtLow, VTName + ".������������(")
		if Pos > 0 then
			Pos2 = instr(Pos, txtLow, ")")
			Pos3 = instr(Pos + 1, txtLow, vbCrLf)
			if Pos2 > Pos3 then
			elseif Pos2 > 0 then
				ColumnParams = Split(mid(txtLow, Pos + len(NewColumnString), Pos2 - Pos - len(NewColumnString)), ",")
				if UBound(ColumnParams) >= 1 then
					if Instr(ColumnParams(0), """") > 0 then
						if Replace(ColumnParams(0), """", "") = NameToSearch then
							if Instr(ColumnParams(1), """") > 0 then
								GetVTColumnType = Replace(ColumnParams(1), """", "")
								exit function
							end if
						end if
					end if
				end if
			end if
			Pos = Pos + 1
		end if
	loop until Pos = 0

end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ����� ��� �� � ��������������� VTName
'
Private Function GetVTFields(VTName, FieldName)
	GetVTFields = "  3"
	NewFieldString = "" + VTName + "[.](������������)[(][""].+[)][;]"
	Set Matches = GetREQuery(txtTrim, NewFieldString)
	for Each NextMatch in Matches
		NextMatchField = Replace( NextMatch.Value, VTName + ".������������(", "",1,-1,1)
		Parameters = Split(NextMatchField,",")
		if UBound(Parameters)>0 Then
			if InStr(1,Parameters(0),"""") > 0 Then ' ������ ���� � ��� �� ���������� � ���������� ���� "KodTov"
				GetVTFields = GetVTFields +  + vbCrLf + picTableField + Replace( Parameters(0), """","")
			End IF
		End IF
	Next
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ������� ��� �� � ��������������� VTName
'� ������ ����� �������� FieldName ���� ��� �� ��� �������
'����� � �������� FieldName �������� ��� �������.
'
Private Function GetVTColumns(VTName, FieldName)
Debug "GetVTColumns - VTName", VTName
Debug "GetVTColumns - FieldName", FieldName

	AllIdentsLine = "|" + VTName + IdentLine
	AllIdents = Split(AllIdentsLine, "|")

	WndType = GetWindType()

	for i = 1 to UBound(AllIdents)
		UnloadString = ".+[.](���������)[(]" + AllIdents(i) + "([,](.*)[)]|[)])"
		Set Matches = GetREQuery(txtLow, UnloadString)
		for Each NextMatch in Matches
			NewIdent = LeftFrom(NextMatch.Value, ".���������(")
			if instr(AllIdentsLine + "|", "|" + NewIdent + "|") = 0 then
				AllIdentsLine = AllIdentsLine + "|" + NewIdent
			end if
		next

		UnloadString = AllIdents(i) + "[.](���������)[(].+[)]"
		Set Matches = GetREQuery(txtLow, UnloadString)
		for Each NextMatch in Matches
			NewIdent = strBetween(NextMatch.Value, ".���������(", ")")
			if instr(AllIdentsLine + "|", "|" + NewIdent + "|") = 0 then
				AllIdentsLine = AllIdentsLine + "|" + NewIdent
			end if
		Next

	next
	AllIdents = Split(AllIdentsLine, "|")

	GetVTColumns = ""
	for i = 1 to UBound(AllIdents)

		SaveI = i

		IdentName = AllIdents(i)
		if Right(LeftFrom(IdentName, "."), 1) = ")" then
			IdentName = GetExprParts(IdentName + ".")
			CanBeFunc = True
		else
			CanBeFunc = False
		end if
		IdentType = GetExprKind(IdentName, 1, CanBeFunc)

		if IdentType = "������" then
			if FieldName = "" then
				QueryFields = Split(Replace(GetQueryAttrs(lCase(IdentName), ""), picAttr, picTableField), vbCrLf)
				for j = 0 to Ubound(QueryFields)
					if instr(GetVTColumns + vbCrLf,  vbCrLf + QueryFields(j) + vbCrLf) = 0 then
						GetVTColumns = GetVTColumns + vbCrLf + QueryFields(j)
					end if
				Next
				exit for
			else
				GetVTColumns = GetQueryAttrs(lCase(IdentName), FieldName)
				Exit Function
			end if
		end if

		sRegIdentName = PrepareForRegularExpression(IdentName)

		NewColumnString = sRegIdentName + "[.](������������|���������������)[(][""].+[)][;]"
		Set Matches = GetREQuery(txtTrim, NewColumnString)
		for Each NextMatch in Matches

			ColumnParamString = StrBetween(NextMatch.Value, "(""", ");")
			ColumnParams = Split(ColumnParamString, ",")
			if UBound(ColumnParams) >= 0 then
				if Instr(ColumnParams(0), """") > 0 then
						ColumnName = Replace(ColumnParams(0), """", "")
						if FieldName = "" then
							if instr(GetVTColumns + vbCrLf,  vbCrLf + ColumnName + vbCrLf) = 0 then
								GetVTColumns = GetVTColumns + vbCrLf + picTableField + ColumnName
							end if
						elseif FieldName = lCase(ColumnName) then
							TypePos = 1
							if instr(lCase(NextMatch.Value), "���������������") > 0 then
								TypePos = 2
							end if
							if UBound(ColumnParams) >= TypePos then
								if instr(ColumnParams(TypePos), """") > 0 then
									GetVTColumns = strBetween(ColumnParams(TypePos), """", """")
								end if
							else
								CrObjLine = "[^�-�A-Z]" + IdentName + "[.]" + ColumnName + "=((�������������|createObject)|����������������������)+[(]+[""]+[�-�A-Z_0-9]+[.]*[�-�A-Z_0-9]*[""]+[)]+[;]+"
								Set CrObjMatches = GetREQuery(txtLow, CrObjLine)
								for each NextM in CrObjMatches
									GetVTColumns = strBetween(NextM.Value, "(""", """)")
								next
							end if
							exit function
						end if
				end if
			end if
		Next

		HasColumns = True
		Pos2 = 0

		UnloadString = ".�����������������������(" + IdentName + ","
		Pos = instr(txtLow, UnloadString)
		if Pos = 0 then
			HasColumns = False
			UnloadString = ".�����������������������(" + IdentName + ")"
			Pos = instr(txtLow, UnloadString)
		end if
		If Pos = 0 and left(wndType, 9) = "��������." then
			UnloadString = "�����������������������(" + IdentName + ","
			Pos2 = instr(txtLow, UnloadString)
			if Pos2 = 0 then
				HasColumns = False
				UnloadString = "�����������������������(" + IdentName + ")"
				Pos2 = instr(txtLow, UnloadString)
			end if
			Pos = Pos2
		end if


		if Pos > 0 then
			if Pos2 = 0 then
				AllStrings = Split(left(txtLow, Pos - 1), vbCrLf)
				DocIdent = AllStrings(UBound(AllStrings))
				if Right(LeftFrom(IdentName, "."), 1) = ")" then
					IdentName = GetExprParts(IdentName + ".")
					CanBeFunc = True
				else
					CanBeFunc = False
				end if
				DocKind = GetExprKind(DocIdent, 1, CanBeFunc)
			else
				DocKind = wndType
			end if
			if instr(DocKind, "��������.") = 1 then
				if FieldName = "" then
					DocKind = rightfrom(docKind, ".")
					AttrColumns = Split(Replace(GetAttrs(2, rightfrom(DocKind, "."), 1, 1), picObjField, picTableField), vbCrLf)
					for j = 0 to UBound(AttrColumns)
						if instr(GetVTColumns + vbCrLf,  vbCrLf + AttrColumns(j) + vbCrLf) = 0 then
							GetVTColumns = GetVTColumns + vbCrLf + AttrColumns(j)
						end if
					Next
				else
					GetVTColumns = GetAttrType(2, rightfrom(DocKind, "."), FieldName, 1)
					exit function
				end if
			end if
		end if

		i = SaveI

	Next

	if GetVTColumns <> "" then GetVTColumns = Mid(GetVTColumns, 3)
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ��� ��������� ��� ����������� � ����� RefType
'��� ������ ������, ���� ���������� �� �������� �����������
'
Private function GetRefOwnerType(RefType)

	GetRefOwnerType = ""

	Set MainObj = MetaData.TaskDef
	Set ObjArray = MainObj.Childs(1)
	for i = 0 to ObjArray.Count - 1
		if LCase(ObjArray(i).Name) = RefType then
			Owner = Trim(ObjArray(i).Props(3))
			if Len(OwnerProp) > 2 then
				GetRefOwnerType = Mid(OwnerProp, 2, len(OwnerProp) - 2)
			end if
			Exit for
		end if
	Next

end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ��������� ������� ����������� ������������ �����
' ObjNumber - ����� ������� � ��������� MetaData.TaskDef.Child
' ObjKind - ������������� ��� ID �������
' ChildsLBound � ChildsUBound ������ � ������� ������ ��������� ����������
' �������� ��� ��������� ���� ���������� ����� � �� ��������� ��� �������� ������ ���� 0 � 1
' ������ ����� - 0
'
Private Function GetAttrs(ObjNumber, ObjKind, ChildsLBound, ChildsUBound)

	GetAttrs= ""

	Set MainObj = MetaData.TaskDef
	Set ObjArray = MainObj.Childs(ObjNumber)

	if left(ObjKind, 1) = "#" then
		ObjID = CLng(Mid(ObjKind, 2))
		for i = 0 to ObjArray.Count - 1
			if ObjArray(i).ID = ObjID then

				Set RefObj = ObjArray(i)
				for j = ChildsLBound to ChildsUbound

					if ObjNumber = 2 and j = 1 then
						pic = picObjTableField
					else
						pic = picObjField
					end if

					Set ObjArray2 = RefObj.Childs(j)
					For i2 = 0 To ObjArray2.Count - 1
						Set AttrObj = ObjArray2(i2)
						GetAttrs = GetAttrs +  vbCrLf + pic + AttrObj.Name
					Next
				Next

				Exit for
			end if
		Next
	else
		for i = 0 to ObjArray.Count - 1
			if LCase(ObjArray(i).Name) = ObjKind then


				Set RefObj = ObjArray(i)
				for j = ChildsLBound to ChildsUbound

					if ObjNumber = 2 and j = 1 then
						pic = picObjTableField
					else
						pic = picObjField
					end if

					Set ObjArray2 = RefObj.Childs(j)
					For i2 = 0 To ObjArray2.Count - 1
						Set AttrObj = ObjArray2(i2)
						GetAttrs = GetAttrs +  vbCrLf + pic + AttrObj.Name
					Next
				Next

				Exit for
			end if
		Next
	end if

	if ObjNumber = 2 And ChildsLBound = 0 then
		Set ObjArray = MainObj.Childs(15)
		for i = 0 to ObjArray.Count - 1
			GetAttrs = GetAttrs + vbCrLf + picCommonDocField + ObjArray(i).Name
		Next
	end if

	GetAttrs = Mid(GetAttrs, 3)

End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ��� ��������� �������
' ObjNumber - ����� ������� � ��������� MetaData.TaskDef.Child
' ObjKind - ������������� �������
' AttrName - ��� ��������
' ChildsUBound - ������� ����� ��������� ����������
'
Private Function GetAttrType(ObjNumber, ObjKind, AttrName, ChildsUBound)

	GetAttrType= ""

	Set MainObj = MetaData.TaskDef
	Set ObjArray = MainObj.Childs(ObjNumber)

	if left(ObjKind, 1) = "#" then
		ObjID = CLng(Mid(ObjKind, 2))
		for i = 0 to ObjArray.Count - 1
			if ObjArray(i).ID = ObjID then

				Set RefObj = ObjArray(i)
				for j = 0 to ChildsUbound
					Set ObjArray2 = RefObj.Childs(j)
					For i2 = 0 To ObjArray2.Count - 1
						Set AttrObj = ObjArray2(i2)
						if lCase(AttrObj.Name) = AttrName then
							GetAttrType = AttrObj.Props(3) + "." + AttrObj.Props(4)
							exit function
						end if
					Next
				Next

				Exit for
			end if
		Next
	else
		for i = 0 to ObjArray.Count - 1
			if lCase(ObjArray(i).Name) = ObjKind then

				Set RefObj = ObjArray(i)
				for j = 0 to ChildsUbound
					Set ObjArray2 = RefObj.Childs(j)
					For i2 = 0 To ObjArray2.Count - 1
						Set AttrObj = ObjArray2(i2)
						if lCase(AttrObj.Name) = AttrName then
							GetAttrType = AttrObj.Props(3) + "." + AttrObj.Props(4)
							exit function
						end if
					Next
				Next

				Exit for
			end if
		Next
	end if


End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'���������� ������ ����� ��������� �� ������ � MetaData.TaskDef.Childs
' ��������� - 0, ����������� - 1, ��������� - 2 � �.�.
'
Private Function GetKinds(ObjNumber, kindView)

	GetKinds = ""
	Set ObjArray = MetaData.TaskDef.Childs(ObjNumber)
	for i = 0 to ObjArray.Count - 1
		if KindView = "" then
			GetKinds = GetKinds + vbCrLf + picMetaDataKind + ObjArray(i).Name
		elseif lCase(ObjArray(i).Name) = KindView then
			GetKinds = ObjArray(i).Name
			Exit Function
		end if
	Next
	GetKinds = Mid(GetKinds, 3)

end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetSubkontoKind(SubkKind)

	GetSubkontoKind = ""
	Set ObjArray = MetaData.TaskDef.Childs(18)
	for i = 0 to ObjArray.Count - 1
		if lCase(ObjArray(i).Name) = SubkKind then
			  GetSubkontoKind = ObjArray(i).Props(3) + "." + ObjArray(i).Props(4)
		end if
	Next

End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Function GetDocFromPath(path)
	Set GetDocFromPath = Nothing

	folder = CommonScripts.FSO.GetBaseName(CommonScripts.FSO.GetParentFolderName(path))
	file = CommonScripts.FSO.GetBaseName(path)

	docType = "����������"
	docKind = file
	docForm = "�����"
	if instr(lCase(path), "\modules\subconto\") > 0 then
	elseif instr(lCase(path), "\modules\subfolder\") > 0 then
		docForm = "�����������"
	elseif instr(lCase(path), "\modules\sublist\") > 0 then
		docForm = "�����������"
		if lCase(folder) <> "sublist" then
			docForm = docForm & "." & file
			docKind = folder
		end if
	elseif instr(lCase(path), "\modules\document\") > 0 then
		docType = "��������"
	elseif instr(lCase(path), "\modules\journal\") > 0 then
		docForm = "�����������"
		if lCase(folder) <> "sublist" then
			docForm = docForm & "." & file
			docKind = folder
		end if
		docType = "������"
	elseif instr(lCase(path), "\modules\report\") > 0 then
		docType = "�����"
	elseif instr(lCase(path), "\modules\calcvar\") > 0 then
		docType = "���������"
	elseif instr(lCase(path), "\modules\transact\") > 0 then
		docType = "��������"
		docForm = "������ ���������"
	else
		exit function
	end if

	'Message docType & "." & docKind & "." & docForm
	Set GetDocFromPath = Documents.item(docType & "." & docKind & "." & docForm)
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ��������� ����� ���������� VarName ����� ���������� �������,
' �������� ����������� �������� �����. ���������� ��� ��� ������ ������
' ���� ������������� �� �������� ���������� �������� �������
'
Private Function SearchInAttrs(doc, VarName)
	SearchInAttrs = ""

	Set lDoc = GetDocFromPath(doc.Path)
	if Not lDoc Is Nothing then
	else
		set lDoc = doc
	end if
	DocNameParts = Split(lcase(lDoc.name), ".")
	If UBound(DocNameParts) > -1 then
		Select Case DocNameParts(0)
		Case "��������":
			SearchInAttrs = GetAttrType(2, DocNameParts(1), lCase(VarName), 1)
		Case "����������":
			SearchInAttrs = GetAttrType(1, DocNameParts(1), lCase(VarName), 0)
		end select
	end if
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���������� ������ ��������������� ������� ������� �����
'
Private Function GetDialogControls()

	GetDialogControls = ""
	Set DocFrm = Windows.ActiveWnd.Document
	If DocFrm = docWorkBook then
		Set DocFrm = DocFrm.Page(0)
	elseif DocFrm = docText then
		'��������c� �������� ����� ��� ������, ������������ � ������� ����������������
		Set DocFrm = GetDocFromPath(DocFrm.Path)
		If DocFrm = docWorkBook then
			Set DocFrm = DocFrm.Page(0)
		else
			exit function
		end if
	else
		exit function
	end if

	DialogTxtStrings = Split(DocFrm.Stream, vbCrLf)

	Level = ""
	for i = 4 to UBound(DialogTxtStrings)

		NextStr = DialogTxtStrings(i)

		if Level = "Table" or Level = "Controls" then

			if Level = "Table" then
				IdentPos = 7
			else
				IdentPos = 12
			end if

			Do while right(NextStr, 2) <> "},"  and i < UBound(DialogTxtStrings)
				i = i + 1
				NextStr = NextStr + vbCrLf + DialogTxtStrings(i)
			loop

			Params = Split(NextStr, """,""")
			if UBound(Params) > 12 then
				if Params(IdentPos) <> "" and Params(IdentPos) <> "0" then
					GetDialogControls = GetDialogControls + vbCrLf + picFormAttr + Params(IdentPos)
				end if
			end if
		end if

		if left(NextStr, 9) = "{""Fixed""," then
			Level = "Table"
		elseif left(NextStr, 12) = "{""Controls""," then
			Level = "Controls"
		end if

	next

	if GetDialogControls <> "" then
		GetDialogControls = mid(GetDialogControls, 3)
	end if

End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' �������� ������ ���� ������� ��� ��������� �����
'
Private Function GetFormsAttrsMethods(doc, AttrName)
	GetFormsAttrsMethods = ""

	b = ReadMethodsFromFile(GetFormsAttrsMethods, "������������", "")
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ��������� ����� ���������� VarName ����� ��������� �����
'
Private Function SearchInForm(doc, VarName)

	SearchInForm = ""

	Set DocFrm = Windows.ActiveWnd.Document
	If DocFrm = docWorkBook then
		Set DocFrm = DocFrm.Page(0)
	else
		exit function
	end if

	DialogTxt = lCase(DocFrm.Stream)

	Pos = instr(DialogTxt, """,""" + lCase(VarName) + """,""")
	if Pos > 0 then
		AllStrings = Split(Left(DialogTxt, Pos - 1), vbCrLf)
		Params = Split(AllStrings(UBound(AllStrings)), """,""")
		if UBound(Params) = 11 then

			if Params(1) = "combobox" or Params(1) = "listbox" then
				SearchInForm = "��������������"
			elseif Params(1) = "table" then
				SearchInForm = "���������������"
			elseif Params(1) = "bmasked" then
				AllStrings = Split(Mid(DialogTxt, Pos + 4), vbCrLf)
				Params = Split(AllStrings(0), """,""")
				if Params(2) = "o" then
					SearchInForm = "��������.#" + Params(5)
				elseif Params(2) = "b" then
					SearchInForm = "����������.#" + Params(5)
				elseif Params(2) = "e" then
					SearchInForm = "������������.#" + Params(5)
				end if
			end if
		end if
	end if
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ����������� ����� ������ StringToAnalize � ���������� ������
' � �������� ��������� ����� � ����� �����
' ��� �������� ��� ������ "��.�������������(�����.��������������()."
' ������� ������ "�����.��������������,
'
Private Function GetExprParts(StringToAnalize)
	inString = False
	PrePos = 0: CurPos = 1
	While CurPos > 0
		CurPos = InStr(CurPos, StringToAnalize, """")
		If CurPos > 0 Then
			PrePos = CurPos
			inString = Not inString
			CurPos = CurPos + 1
		End If
	Wend
	If inString Then
		GEP = Split(Mid(StringToAnalize, PrePos + 1), ".")
	Else
		GEP = Split(StringToAnalize, ".")
	End If

	BreakingSymbs = " ;/-:',=+*|" + vbTab
	breakingPart = LBound(GEP)
	StringToAnalize = GEP(UBound(GEP))

	If UBound(GEP) = 0 Then Q = 0 Else Q = 1

	BlocksVal = 0
	For i = UBound(GEP) - Q To LBound(GEP) Step -1
		If Right(GEP(i), 1) = "(" Then
			GEP(i) = Left(GEP(i), Len(GEP(i)) - 1)
		End If
		if BlocksVal > 0 then
			BlocksVal = 0
		end if
		TextFlag = False
		For i1 = Len(GEP(i)) To 1 Step -1
			If Mid(GEP(i), i1, 1) = """" Then
				TextFlag = Not TextFlag
			ElseIf TextFlag = True Then
			ElseIf InStr(BreakingSymbs, Mid(GEP(i), i1, 1)) and BlocksVal >= 0 Then
				GEP(i) = Mid(GEP(i), i1 + 1)
				breakingPart = -i
				Exit For
			ElseIf Mid(GEP(i), i1, 1) = "(" Then
				Blocksval = Blocksval + 1
			ElseIf Mid(GEP(i), i1, 1) = ")" Then
				Blocksval = Blocksval - 1
			End If
		Next
		If TextFlag And i > LBound(GEP) Then
			GEP(i - 1) = GEP(i - 1) + "." + GEP(i)
			Q = Q + 1
		ElseIf BlocksVal < 0 then
			Q = Q + 1
		Else
	  	If Blocksval > 0 Then
			While Blocksval > 0
			   GEP(i) = RightFrom(GEP(i), "(")
			   Blocksval = Blocksval - 1
			Wend
			breakingPart = -i
	  	End If
	  	If InStr(GEP(i), "(") > 0 Then
			GEP(i) = LeftFrom(GEP(i), "(")
	  	End If
	  	If breakingPart < 0 Then
			breakingPart = -breakingPart
			Exit For
	  	End If
		End If
	Next

	GetExprParts = ""
	For i = breakingPart To UBound(GEP) - Q
		GetExprParts = GetExprParts + GEP(i) + "."
	Next
	If Len(GetExprParts) > 1 Then
		GetExprParts = Left(GetExprParts, Len(GetExprParts) - 1)
	End If
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���������� ��������� �������� �� ����� ��������� � ������ ��������.
' �������� ����������� � ����� ������������ ����� ���� ��� ��������, ��������������
' � �.�. ������� � ����������� �� ���� ������ ����� ����� ��� ��� ���� �����.
'
Private Function CheckModuleType(docType, CheckKind)
	CheckModuleType = false
	Select Case CheckKind
	Case "1": CheckModuleType = (instr(doctype, ".�����") > 0 or instr(doctype, ".ert") > 0 or instr(doctype, "cwbmoduledoc") > 0 _
		or instr(docType, "\modules\document\") > 0 _
		or instr(docType, "\modules\journal\") > 0 _
		or instr(docType, "\modules\subconto\") > 0 _
		or instr(docType, "\modules\subfolder\") > 0 _
		or instr(docType, "\modules\sublist\") > 0 _
		or instr(docType, "\modules\report\") > 0 _
		or instr(docType, "\modules\calcvar\") > 0)
	Case "2": CheckModuleType = (right(doctype, 16) = "������ ���������")
	Case "3": CheckModuleType = (left(doctype, 11) = "����������.")
	Case "4": CheckModuleType = (left(doctype, 9) = "��������." or left(doctype, 7) = "������.")
	End select
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetParamName(ProcName, VarName, ParamNumber)
	GetParamName = ""

	ProcString = ProcName + "\((.+)\)"
	Set Matches = GetREQuery(txtLow, ProcString)
	for Each NextMatch in Matches
		Params = Split(NextMatch.Submatches(0), ",")
		if UBound(Params) >= ParamNumber then
			Param = replace(Params(ParamNumber), vbCrLf, "")
			if Param <> VarName then
				GetParamName = Param
				Exit function
			end if
		end if
	next
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Function FindInTypeDefs(VarName, TypeString)
	Dim ElseLevels(100)

	ModuleLines = Split(txtLow, vbCrLf)
	FindInTypeDefs = ""

	for j = 1 to 2
		if j = 1 and TypeString <> "" then
		else
			sRegVarNameSearch = PrepareForRegularExpression(VarName)

			if j = 1 then
				TypeDefLine = "��������������\(" + sRegVarNameSearch + "\)=""([\w�-��]+)"""
				stringToFind = "������������������(" + lcase(VarName) + ")="""
			else
				TypeDefLine = "[^\w�-��]*" + sRegVarNameSearch + ".���\(\)=""([\w�-��]+)"""
				stringToFind = "����" + lcase(VarName) + ".���()="""
			end if

			Set DefMatches = GetREQuery(txtLow, TypeDefLine)

			if DefMatches.Count > 0 then
				CurrentLevel = 0
				for i = doc.SelStartLine to 0 step -1
					if Instr(ModuleLines(i), stringToFind) > 0 then
						if CurrentLevel >= 0 then
							FindInTypeDefs = strBetween(ModuleLines(i), ")=""", """")
							If TypeString <> "" then
								FindInTypeDefs = TypeString + "." + FindInTypeDefs
							elseif j = 2 then
								for i1 = 0 to 1
									Set ObjArray = MetaData.TaskDef.Childs(i1 + 1)
									for i2 = 0 to ObjArray.Count - 1
										if lCase(ObjArray(i2).Name) = FindInTypeDefs then
											FindInTypeDefs = Trim(Mid("����������.��������.  ", i1 * 11 + 1, 11)) + FindInTypeDefs
											exit function
										end if
									Next
								next
							end if
						end if
						exit for
					elseif Instr(ModuleLines(i), "���������") = 1 then
						CurrentLevel = CurrentLevel - 1
					elseif Instr(ModuleLines(i), "���������") = 1 then
						if ElseLevels(CurrentLevel + 10) = 0 then
							CurrentLevel = CurrentLevel - 1
							ElseLevels(CurrentLevel + 10) = 1
						end if
					elseif Instr(ModuleLines(i), "����") = 1 then
						ElseLevels(CurrentLevel + 10) = 0
						CurrentLevel = CurrentLevel + 1
						LastElse = 0
					elseif Instr(ModuleLines(i), "����������") = 1 then
						CurrentLevel = CurrentLevel - 1
					elseif Instr(ModuleLines(i), "����") = 1 then
						CurrentLevel = CurrentLevel + 1
						LastElse = 0
					elseif Instr(ModuleLines(i), "���������") = 1 then
						exit for
					elseif Instr(ModuleLines(i), "�������") = 1 then
						exit for
					end if

					if doc.SelStartLine - i = 50 then
						Exit for
					end if
				Next
			End if
		end if

		if j=1 then
			for i = 0 to 30
				ElseLevels(i) = 0
			next
		end if
	Next

	if FindInTypeDefs = "" then FindInTypeDefs = TypeString
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
' ����������� ������ ��� ������������� � ��������
' �������� ����.������� �������� �� ��� �� � ����������� ����� \
' �.�. �������� "(" ������ �� "\("
'
Private Function PrepareForRegularExpression(SourceString)
	DestString = SourceString

	DestString = Replace(DestString, "(", "\(")
	DestString = Replace(DestString, ")", "\)")
	DestString = Replace(DestString, ".", "\.")
	DestString = Replace(DestString, "|", "\|")
	DestString = Replace(DestString, "[", "\[")
	DestString = Replace(DestString, "]", "\]")

	PrepareForRegularExpression = DestString

	if SourceString <> DestString then
		Debug	"��������� � ���������� ��������� <"+SourceString+">", DestString
	end if

End Function ' PrepareForRegularExpression

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ���������� ��� ���������� VarName
function GetMainType(VarName, CountCalls)
ProfilerEnter("GetMainType 1")
	GetMainType =""
	RepeatCount = 0
	VarNameSearch = VarName
	Debug	"��� ��� ��� <"+VarNameSearch+">", VarNameSearch

	'�������� �� �������������
	'��� ����� ������� ���� ���������� ��� ������ ����������� ����
	'	//:<����������>
	'����� ����� ���������� (kuntashov@yandex.ru, 2005-06-08)
	'CrObjLine = "[^\w�-��]" + VarNameSearch + "(=(�������������|createobject|����������������������)\(""([\w�-��\d_.]+)""\);*|[,;=]*//:([\w�-��\d_.]+))"

	' - ������� ������� ����� ������������� - ( artbear 2005-06-09 )
	sRegVarNameSearch = PrepareForRegularExpression(VarNameSearch)

	'CrObjLine = "[^\w�-��]" + "("+sRegVarNameSearch +"|"+ "�����"+ sRegVarNameSearch+")" + "(=(�������������|createobject|����������������������)\(""([\w�-��\d_.]+)""\);*|[,;=]*//:([\w�-��\d_.]+))"
	'CrObjLine = "[^\w�-��]*" + "("+sRegVarNameSearch +"|"+ "�����"+ sRegVarNameSearch+")" + "(=(�������������|createobject|����������������������)\(""([\w�-��\d_.]+)""\);*|[,;=]*//:([\w�-��\d_.]+))"
	CrObjLine = "[^\w�-��]*" + "("+sRegVarNameSearch +"|"+ "�����"+ sRegVarNameSearch+")" + "(=[\w�-��\d_.]*\.?(�������������|�����������������|wrapactivex|createobject|����������������������)\(""([\w�-��\d_.]+)""\);*|[,;=]*//:([\w�-��\d_.]+))"	
	
ProfilerEnter("GetMainType 2")
	Set CrObjMatches = GetREQuery(txtLow, CrObjLine)
ProfilerLeave("GetMainType 2")
	for each NextMatch in CrObjMatches
		GetMainType = NextMatch.Submatches(3)

		if GetMainType <> "" then 		Debug	"����� ���", GetMainType

		' trdm 2005/10/11
		MethodCreation = CrObjMatches.Item(0).Submatches(2)
		if (MethodCreation = "createobject") or (MethodCreation = "�������������") Then
			'On Error Resume Next
			if CommonScripts.GetScriptByName("intsOLEGenerator") is Nothing then
			else
				if Scripts("intsOLEGenerator").IntsGenerator( GetMainType ) Then
					ReloadDictionary
				End If
			End If
			'On Error goto 0
		end if
		' end trdm 2005/10/11

		if GetMainType = "" then
			GetMainType = NextMatch.Submatches(4)

			if GetMainType <> "" then 	Debug	"����� ���", GetMainType

		end if
		TypeDef_Pos = NextMatch.FirstIndex

	' TODO - ���-�� ����� ������������ � �������� ���������, ���� ������ ������� ��������� ����������� �� ������ (�.� ����� ���� ��������)

	next

	if GetMainType <> "" then
		if OtherTypesDefine.Exists(GetMainType) then
			GetMainType = OtherTypesDefine.Item(GetMainType)
		end if
		if left(GetMainType, 3) = "���" or left(GetMainType, 3) = "���" then
			CreateObjectFlag = "1"
		end if
	elseif MainTypes.Exists(LCase(VarNameSearch)) then
		GetMainType = MainTypes.Item(LCase(VarNameSearch))
		Pos = instr(GetMainType, "#")
		if Pos > 0 then
			if CheckModuleType(lcase(Doc.name), rightfrom(GetMainType, "#")) then
				GetMainType = leftFrom(GetMainType, "#")
				if GetMainType = "FRMTYPE" then
					GetMainType = GetWindType()
					if left(GetMainType, 7) = "������." then
						GetMainType = "��������"
					end if
				end if
			else
				GetMainType = ""
			end if
		end if
	end if

	if GetMainType = "" then
		GetMainType = lcase(SearchInAttrs(doc, VarNameSearch))
	end if

	if GetMainType = "" then
		GetMainType = lCase(SearchInForm(doc, VarNameSearch))
	end if

	if GetMainType = "" then

		CrObjLine = "[^\w�-��]" + sRegVarNameSearch + "=(.+);"
		NewSearch = ""
		Set CrObjMatches = GetREQuery(txtLow, CrObjLine)
		for each NextMatch in CrObjMatches
			NewSearch = NextMatch.Submatches(0)
		next

		if NewSearch <> "" then
			IdentLine = IdentLine + "|" + NewSearch
			if Right(LeftFrom(NewSearch, "."), 1) = ")" then
				CanBeFunc = True
			else
				CanBeFunc = False
			end if
			GetMainType = GetExprKind(GetExprParts(NewSearch + "."), CountCalls + 1, CanBeFunc)
		end if
	end if

	if GetMainType = "" then
		ModuleStrings = Split(txtLow, vbCrLf)
		for i = doc.SelStartLine - 1 to 0 step -1
			if Left(ModuleStrings(i), 7) = "�������" Or _
			   Left(ModuleStrings(i), 9) = "���������" then

			   	Pos1 = instr(ModuleStrings(i), "(")
			   	Pos2 = instr(ModuleStrings(i), ")")

			   	if Pos1 = 0 or Pos2 = 0 or Pos1 > Pos2 then
			   		exit for
			   	end if

			   	Params = split(mid(ModuleStrings(i), Pos1 + 1, Pos2 - Pos1 - 1), ",")
			   	for j = 0 to UBound(Params)
			   		if LeftFrom(Params(j), "=") = VarNameSearch then

			   			if Left(ModuleStrings(i), 7) = "�������" then
			   				ProcName = Mid(ModuleStrings(i), 8, Pos1 - 8)
			   			else
			   				ProcName = Mid(ModuleStrings(i), 10, Pos1 - 10)
			   			end if
			   			NewSearch = GetParamName(ProcName, VarNameSearch, j)
			   			if NewSearch <> "" then
			   				IdentLine = IdentLine + "|" + NewSearch
			   				GetMainType = GetExprKind(GetExprParts(NewSearch + "."), CountCalls + 1, False)
			   			end if
			   			exit for
			   		end if
			   	next
			   	exit for
			end if
		next
	end if

	if GetMainType = "" then
		GetMainType = FindInTypeDefs(VarNameSearch, "")
	elseif GetMainType = "��������." or GetMainType = "��������" then
		GetMainType = FindInTypeDefs(VarNameSearch, "��������")
	elseif GetMainType = "����������." or GetMainType = "����������" then
		GetMainType = FindInTypeDefs(VarNameSearch, "����������")
	end if
ProfilerLeave("GetMainType 1")
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function FindInFunc(FunctionName, CountCalls)
	FindInFunc = ""

	FindStr = "(�������" + FunctionName + "[(]|������������)"
	Set FuncMatches = GetREQuery(txtLow, FindStr)
	StartPos = 0 : EndPos = 0
	for each NextMatch in FuncMatches
		if StartPos = 0 then
			if instr(NextMatch.Value, "�������") = 1 then
				StartPos = NextMatch.FirstIndex
			end if
		elseif instr(NextMatch.Value, "�����") = 1 then
			EndPos = NextMatch.FirstIndex
			Exit for
		end if
	next

	if StartPos > 0 and EndPos > StartPos then
		FuncText = Mid(txtLow, StartPos, EndPos - StartPos)
		FindStr = "(�������).+"
		Set ReturnMatches = GetREQuery(FuncText, FindStr)
		For Each NextMatch in ReturnMatches
			IdentName = leftFrom(Mid(NextMatch.Value, 8), ";")
			IdentName = leftFrom(IdentName, chr(13))
			FindInFunc = GetExprKind(GetExprParts(IdentName + "."), CountCalls + 1, False)
			if FindInFunc <> "" then
				Exit function
			end if
		Next
	End if
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'		�������� ��� ���������
'		ConstantKind = ��� ���������, ��������, �����������
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Function GetConstantType(ConstantKind)
  GetConstantType = ""

	Set MainObj = MetaData.TaskDef
	Set ObjArray = MainObj.Childs(0)

		for i = 0 to ObjArray.Count - 1
			if lCase(ObjArray(i).Name) = ConstantKind then

				Set RefObj = ObjArray(i)
				GetConstantType = LCase(RefObj.Type.FullName)
				Exit for
			end if
		Next

End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetExprKind(CurrentLine, CountCalls, CanBeFunc)
	GetExprKind = ""
	TypeDef_Pos = -1
	if CountCalls > 10 then
		Exit function
	end if

	ExprParts = Split(CurrentLine, ".")
	IdentLine = ""

	ObjType = ""
	for i = 0 to UBound(ExprParts)
		SaveI = i
		if i = 0 then
'ProfilerEnter("GetExprKind 1")
			ObjType = GetMainType(ExprParts(0), CountCalls)
'ProfilerLeave("GetExprKind 1")
			if ObjType = "" And CanBeFunc = True then
ProfilerEnter("GetExprKind 2")
				ObjType = FindInFunc(ExprParts(0), CountCalls)
ProfilerLeave("GetExprKind 2")
			end if
		else
			CreateObjectFlag = "0"
			ObjType = ""

			DefInfo = "_"
			if TypesStructureExt.Exists(ObjParts(0) + "." + ExprParts(i)) then
				DefInfo = TypesStructureExt.Item(ObjParts(0) + "." + ExprParts(i))
			elseif TypesStructure.Exists(ObjParts(0)) then
				DefInfo = TypesStructure.Item(ObjParts(0))
			end if

			DefInfoParts = Split(DefInfo, "|")

			Select case DefInfoParts(0)
			case "VALUE":
				ObjType = DefInfoParts(1)
			case "KINDNEW":
				ObjType = DefInfoParts(1) + "." + ExprParts(i)
			case "KINDOLD":
				if UBound(ObjParts) = 0 then
					ObjType = DefInfoParts(1)
				else
					ObjType = DefInfoParts(1) + "." + ObjParts(1)
				end if
			case "ATTR":
				if UBound(ObjParts) > 0 then
					ObjType = GetAttrType(CLng(DefInfoParts(1)), ObjParts(1), ExprParts(i), CLng(DefInfoParts(2)))
				end if
			case "REFOWNER"
				if UBound(ObjParts) > 0 then
					ObjType = GetRefOwnerType(ObjParts(1))
				end if
			case "VTCOLUMN"
				ObjType = GetVTColumns(lcase(ExprParts(0)), lCase(ExprParts(i)))
			case "QUERYATTR"
				ObjType = GetQueryAttrs(lcase(ExprParts(0)), lCase(ExprParts(i)))
			case "SUBKREF"
				ObjType = GetSubkontoKind(lCase(ExprParts(i)))
			end select
		end if

		if ObjType = "" then
			exit function
		else
			ObjType = lCase(ObjType)
			ObjParts = Split(ObjType, ".")
		end if

		i = SaveI
	next

	' ����� ������ �������� ?
	if not IsEmpty(ObjParts) then

		' ��� ��������� ���������.����������� ������ ������������ �������� ��� ����
		if ObjParts(0) = "���������" then
				ObjType = GetConstantType(ObjParts(1))
		end if ' ���������

	end if

	GetExprKind = ObjType
End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetWithVarName()
	GetWithVarName = "."

	ModuleStrings = Split(txtTrim, vbCrLf)
	for i = doc.SelStartLine - 1 to 0 step -1
		Pos = instr(ModuleStrings(i), "//_�_")

		PosF = instr(lCase(ModuleStrings(i)), "���������") + instr(lCase(ModuleStrings(i)), "�������")
		if Pos > 0 then
		   	GetWithVarName = Trim(Mid(ModuleStrings(i), Pos + 5))
		   	Exit Function
		elseif PosF = 1 then
			Exit Function
		end if
	next
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function Get1CppIntsFolder()
	Get1CppIntsFolder = DataDir & "\1�++"
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ������� ���� � ����� Intell, � ��� ����� � ����� Intell\1�++
Private Function GetIntsFilePath(FileNameWithoutExtension)
	Dim FileName
  GetIntsFilePath = ""

	FileName = DataDir & "\" & FileNameWithoutExtension & ".ints"
	if not FSO.FileExists(FileName) then FileName = Get1CppIntsFolder & "\" & FileNameWithoutExtension + ".ints"
	if not FSO.FileExists(FileName) then Exit Function

  GetIntsFilePath = FileName
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

'::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function ReadMethodsFromFile(Methods, FileNm, LoadedString)
	ReadMethodsFromFile = true
	if Instr(LoadedString, "," + lCase(FileNm) + ",") > 0 then exit function

	ReadMethodsFromFile = false

	FileName = GetIntsFilePath(FileNm)
	if FileName = "" then
			Exit Function
	end if

	Set FileStream = fso.OpenTextFile(FileName, 1, false)
	if FileStream.AtEndOfStream = true then  Exit Function

	AllMeth = FileStream.ReadAll()
	AllMethods = Split(AllMeth, vbCrLf)

	Dividor = ""
	If Methods <> "" Then Dividor = vbCrLf
	for i = 0 to UBound(AllMethods)
		if Mid(AllMethods(i),1,1) <= CreateObjectFlag then
			if instr(AllMethods(i), ")") > 0 then
				Methods = Methods + Dividor + picMeth + Mid(AllMethods(i), 6)
			else
				Methods = Methods + Dividor + picAttr + Mid(AllMethods(i), 6)
			end if
			Dividor = vbCrLf
		end if
	next

	ReadMethodsFromFile = true
	LoadedString = LoadedString + lCase(FileNm) + ","
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function FindByKeyMeths(VarName)

	FindByKeyMeths = ""

	IdentStr = "([^�-�A-Z]|��)" + VarName + "[.][�-�A-Z]+[(]"
	Set AllMatches =	GetREQuery(txtTrim, IdentStr)
	for each NextMatch in AllMatches
		MethodName = Replace(RightFrom(NextMatch.Value, "."), "(", "")
		if KeyMeths.Exists(MethodName) then
			FindByKeyMeths = KeyMeths.Item(MethodName)
			exit for
		end if
	Next

end function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetMethodsList(CurrentLine)
	GetMethodsList = ""
	CreateObjectFlag = "0"
	TypeDef_Pos = -1

	if Right(CurrentLine, 1) = ")" then
		IsMeth = True
	else
		IsMeth = False
	end if

ProfilerEnter("GetMethodsList 1")
	CurrentLn = GetExprParts(CurrentLine + ".")
ProfilerLeave("GetMethodsList 1")
ProfilerEnter("GetMethodsList 2")
	if CurrentLn = "." then
		CurrentLn = GetWithVarName()
		SymbBefore = CurrentLn + SymbBefore
		CurrentLn = lCase(CurrentLn)
		if CurrentLn = "." then
			exit function
		end if
	end if
ProfilerLeave("GetMethodsList 2")

	ExprParts = Split(CurrentLn, ".")

	if instr(lCase(CurrentLine), ExprParts(0) + "(") > 0 then
		CanBeFunc = True
	else
		CanBeFunc = False
	end if

ProfilerEnter("GetMethodsList 3")
	ObjType = GetExprKind(CurrentLn, 1, CanBeFunc)
ProfilerLeave("GetMethodsList 3")

ProfilerEnter("GetMethodsList 4")
	if ObjType = "" then ObjType = FindInTypeDefs(CurrentLn, "")
ProfilerLeave("GetMethodsList 4")
ProfilerEnter("GetMethodsList 5")
	if ObjType = "" then ObjType = FindByKeyMeths(CurrentLn)
ProfilerLeave("GetMethodsList 5")
	if ObjType = "" then exit function

	FinalObjType = ObjType

	ObjParts = Split(ObjType, ".")

	BhvInfo = "_"
	if MethodsBehaviour.Exists(ObjParts(0)) then
		BhvInfo = MethodsBehaviour.Item(ObjParts(0))
	end if
	BhvInfoParts = Split(BhvInfo, "|")

	Methods = ""
	Select Case BhvInfoParts(0)
	Case "ATTR":
		if UBound(ObjParts) > 0 then
			Methods = GetAttrs(CLng(BhvInfoParts(1)), ObjParts(1), 0, CLng(BhvInfoParts(2)))
		end if
	Case "KINDS":
		Methods = GetKinds(CLng(BhvInfoParts(1)), "")
	Case "VTCOLUMNS":
		Methods = GetVTColumns(lCase(ExprParts(0)), "")
	Case "VTFIELDS":
		Methods = GetVTFields(lCase(ExprParts(0)), "")

	Case "DLGCONTROLS":
		Methods = GetDialogControls()
	Case "QUERYATTRS":
		Methods = GetQueryAttrs(lCase(ExprParts(0)), "")

	Case "������������":
		Methods = GetFormsAttrsMethods(doc, lCase(ExprParts(1)))

	Case "_": ' ��������� ������� 1�++ (artbear)
		ObjParts = Split(ObjType, vbCRLF)
	End Select

	if left(ObjParts(0), 3) = "v77" then ObjParts(0) = "v77_Application"

ProfilerEnter("GetMethodsList 6")
	ReadMethodsFromFile Methods, ObjParts(0), ","
ProfilerLeave("GetMethodsList 6")
	if Left(Methods, 2) = vbCrLf then
		Methods = Mid(Methods, 3)
	end if

	GetMethodsList = replace(Methods, vbCrLf + vbCrLf, vbCrLf)
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function QueryOnParamsTip(CurrentLine, ParamNumber)

	QueryOnParamsTip = 0
	Commas = 0

	CurString = CurrentLine
	inString = False
	For i = Len(CurString) To 1 Step -1
		theSymb = Mid(CurString, i, 1)
		If theSymb = """" Then
			inString = Not inString
		ElseIf Not inString Then
			If theSymb = "(" Then
				Pos = Pos - 1
				If Pos < 0 Then
					Pos = -i
					Exit For
				End If
			ElseIf theSymb = ")" Then
				Pos = Pos + 1
			ElseIf theSymb = "," Then
				Commas = Commas + 1
			End If
		End If
	Next
	If Pos < 0 Then
		QueryOnParamsTip = -Pos
		ParamNumber = Commas + 1
	end if

End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function GetListParamValue(CurrentLine)

	GetListParamValue = ""

	ParamNumber = 0
	MethodStartPos = QueryOnParamsTip(CurrentLine, ParamNumber)
	CurrentLn = CurrentLine
	If MethodStartPos > 0 Then CurrentLn = GetExprParts(Left(CurrentLine, MethodStartPos - 1) + ".")


	MethodName = ""
	while Len(CurrentLn) > 0 and Right(CurrentLn, 1) <> "."
		MethodName = Right(CurrentLn, 1) + MethodName
		CurrentLn = Left(CurrentLn, Len(CurrentLn) - 1)
	wend

	if CurrentLn = "" then exit function
	if MethodName = "" then exit function

	CurrentLn = Left(CurrentLn, Len(CurrentLn) - 1)
	if instr(lCase(CurrentLine), leftFrom(CurrentLn, ".") + "(") > 0 then
		CanBeFunc = True
	else
		CanBeFunc = False
	end if

	ExprType = GetExprKind(CurrentLn, 1, CanBeFunc)
	if ExprType = "" then exit function

	ExprParts = Split(ExprType, ".")
	MethQuery = ExprParts(0) + "." + MethodName + "." + CStr(ParamNumber)

	BhvInfo = "_"
	if MethodsParamType.Exists(MethQuery) then
		BhvInfo = MethodsParamType.Item(MethQuery)
	end if

	BhvInfoParts = Split(BhvInfo, "|")
	Select Case BhvInfoParts(0)
	Case "DOCTABLE":
		GetListParamValue = GetAttrs(2, ExprParts(1), 1, 1)
	Case "VTCOLUMNS":
		GetListParamValue = GetVTColumns(LeftFrom(CurrentLn, "."), "")
	Case "KINDS":
		GetListParamValue = GetKinds(CLng(BhvInfoParts(1)), "")
	Case "VTCOLUMNSCOLL"
		GetListParamValue = GetVTColumns(LeftFrom(CurrentLn, "."), "")
		DlgType = 1
	end select


End function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Function GetEqualValues(CurrentLine)
	GetEqualValues = ""

	CurrentLn = GetExprParts(Trim(Mid(CurrentLine, 1, len(CurrentLine) - 1)) + ".")
	LeftVarType = Split(GetExprKind(lcase(CurrentLn), 1, False), ".")

	if UBound(LeftVarType) < 0 then exit function

	Select Case LeftVarType(0)
	Case "�������������������"
		GetEqualValues = GetDialogControls()
	Case "��������������������"
		GetEqualValues = GetKinds(1, "")
	Case "������������������"
		GetEqualValues = GetKinds(2, "")
	Case "������������"
		if UBound(LeftVarType) > 0 then
			GetEqualValues = GetAttrs(4, LeftVarType(1), 0, 0)
			SymbBefore = "������������." + GetKinds(4, LeftVarType(1)) + "."
			SymbAfter = ""
		end if
	Case "��������������������"
		if UBound(LeftVarType) > 0 then
			GetEqualValues = GetAttrs(4, LeftVarType(1), 0, 0)
		end if
	Case "������������"
		GetEqualValues = GetAllPossibleTypesList()
	End select
end function


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ������� �������� ��������� ��� ��������� Expr � ��������� ��������� aDoc
' � ������ ������ ���������� ������ � �����, ���� ��� ���������� �� ������� ���������� ������ ������
Function GetExprType(aDoc, Expr)
	Set Doc = aDoc
	txtTrim = Replace(doc.text, " ", "")
	txtTrim = Replace(txtTrim, vbTab, "")
	txtLow = LCase(txtTrim)
	txtStrings = split(doc.text, vbCrLf)
	txtStrings = split(doc.text, vbCrLf)

	GetExprType = GetExprKind(Expr, 1, false)
End Function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' �������� ' artbear
' � ������� "������ ����" ����� ������ ���������, ������� �������� � ����������� ��������
'
Sub CopyLine()
	doc = ""
  	if Not CheckWindow(doc) then Exit sub

	txtStrings = split(doc.text, vbCrLf)
  	CurrentLine = left(txtStrings(doc.SelStartLine), doc.SelStartCol)
  	SymbolsAfter = Mid(txtStrings(doc.SelStartLine), doc.SelStartCol + 1)

  	doc.range(doc.SelStartLine,doc.SelStartCol + Len(SymbolsAfter), doc.SelEndLine, doc.SelEndCol + Len(SymbolsAfter)) = vbCrLf + CurrentLine
  	doc.MoveCaret doc.SelStartLine + 1, doc.SelStartCol
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' �� ����� ' artbear
'
Private Sub GetType()
	doc = ""
	if Not CheckWindow(doc) then Exit sub

	PrepareModuleText doc.SelStartLine, doc.SelStartCol

	ObjType = GetExprKind(CurrentLine, 1, False)
	if ObjType = "" then
		msgbox "�� ���������"
	else
		msgbox ObjType
	end if
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Sub VsePredopredelennye()
	Dim AllProcs(100)

	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)

	AllProcs(1)  = "����������|�����"
	AllProcs(2)  = "�����������|"
	AllProcs(3)  = "��������������������|"
	AllProcs(4)  = "�����������������������|���������������������, ����"
	AllProcs(5)  = "�����������������|�������������, ����������������"
	AllProcs(6)  = "�����������������������|�������, ���������������������, ����"
	AllProcs(7)  = "���������������|"
	AllProcs(8)  = "���������������|������"
	AllProcs(9)  = "��������������|"
	AllProcs(10) = "����������������������������|"
	AllProcs(11) = "��������������������������������|�������, ������"
	AllProcs(12) = "�����������������������������|"
	AllProcs(13) = "��������������������������������|���������������"
	AllProcs(14) = "�����������������|���������������"
	AllProcs(15) = "������������������������|��������"
	AllProcs(16) = "�����������������|"
	AllProcs(17) = "������������������|���������������"
	AllProcs(18) = "���������������������|�������, �������"
	AllProcs(19) = "����������������|"
	AllProcs(20) = "�����������|"
	AllProcs(21) = "���������|"
	AllProcs(22) = "����������������|�����������������, ���������"
	AllProcs(23) = "����������������������|�������������������, ������������, ������������, �����������"
	AllProcs(24) = "#"

	ProcList = ""
	for i = 1 to 100
		if AllProcs(i) = "#" then exit for

		ProcList = ProcList + vbCrLf + picProcs + leftfrom(AllProcs(i), "|")
	next

	DLLKIND = GetIniParam("CHOICELIST_KIND")
	ProcName = SelectFromList(ProcList, "")

	if ProcName = "" then
		Exit sub
	end if

	if ProcName = "/" then
		doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = "//"
		doc.MoveCaret doc.selStartLine, 2
		Exit sub
	End if

	Txt = "//" + string(60, "=") +  TabInd + "��������� " + ProcName + "("
	BufArr = Split(left(ProcList, instr(ProcList, ProcName)), vbCrLf)
	Pos = UBound(BufArr)

	Txt = Txt + rightfrom(AllProcs(Pos), "|") + ")"
	Txt = Txt + TabInd + vbTab
	Txt = Txt + TabInd + "��������������"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	doc.MoveCaret doc.selStartLine + 2, 1
end sub
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Private Function AutoFillText(ExtAction, CurrentLine, TabInd)
	AutoFillText = ""
	ObjParts = Split(FinalObjType, ".")

	Select Case lCase(ExtAction)
	Case "�������������� ���"
		if UBound(ObjParts) > 0 then
			HeadFields = Split("___��������" + VbCrLf + "___�������" + vbCrLf + GetAttrs(2, ObjParts(1), 0, 0), vbCrLf)
			TableFields = Split(GetAttrs(2, ObjParts(1), 1, 1), vbCrLf)

			CaretMark = "!"
			for i = 0 to UBound(HeadFields)
				AutoFillText = AutoFillText + TabInd + CurrentLine + "." + Mid(HeadFields(i), 4) + " = " + CaretMark + ";"
				CaretMark = ""
			Next

			if UBound(TableFields) >= 0 then
				AutoFillText = AutoFillText + TabInd + "// ���� �� �������"
				AutoFillText = AutoFillText + TabInd + vbTab + CurrentLine + ".�����������();"
				for i = 0 to UBound(TableFields)
					AutoFillText = AutoFillText + TabInd + vbTab + CurrentLine + "." + Mid(TableFields(i), 4) + " = " + CaretMark + ";"
					CaretMark = ""
				Next
				AutoFillText = AutoFillText + TabInd + "//����������;"
			end if
			AutoFillText = AutoFillText + TabInd + CurrentLine + ".��������();"
		end if

	Case "�������������� ����"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".�����.���� = !;"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".�����.��������(1, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".�����.��������(2, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".�����.��������(3, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".������.���� = ;"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".������.��������(1, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".������.��������(2, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".������.��������(3, );"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".����� = ;"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".���������� = ;"
		AutoFillText = AutoFillText + TabInd + CurrentLine + ".������������������ = ;"

	Case "�������������� ���"
		if UBound(ObjParts) > 0 then
			RefFields = Split("___���" + VbCrLf + "___������������" + vbCrLf + GetAttrs(1, ObjParts(1), 0, 0), vbCrLf)

			CaretMark = "!"
			for i = 0 to UBound(RefFields)
				AutoFillText = AutoFillText + TabInd + CurrentLine + "." + Mid(RefFields(i), 4) + " = " + CaretMark + ";"
				CaretMark = ""
			Next
			AutoFillText = AutoFillText + TabInd + CurrentLine + ".��������();"
		end if

	Case "�������������� ��"
		if UBound(ObjParts) > 0 then
			TableFields = Split(GetAttrs(2, ObjParts(1), 1, 1), vbCrLf)
			CaretMark = "!"
			for i = 0 to UBound(TableFields)
				AutoFillText = AutoFillText + TabInd + CurrentLine + "." + Mid(TableFields(i), 4) + " = " + CaretMark + ";"
				CaretMark = ""
			Next
		end if

	Case "�������������� ��"
		TableFields = Split(GetVTColumns(lCase(CurrentLine), ""), vbCrLf)
		CaretMark = "!"
		for i = 0 to UBound(TableFields)
			AutoFillText = AutoFillText + TabInd + CurrentLine + "." + Mid(TableFields(i), 4) + " = " + CaretMark + ";"
			CaretMark = ""
		Next

	Case "����"
		if ObjParts(0) = "��������" then
			AutoFillText = AutoFillText + TabInd + "���� " + CurrentLine + ".����������������() = 1 ����"
			AutoFillText = AutoFillText + TabInd + vbTab + "!" + TabInd + "����������;"
		elseif ObjParts(0) = "����������" then
			AutoFillText = AutoFillText + TabInd + "���� " + CurrentLine + ".���������������() = 1 ����"
			AutoFillText = AutoFillText + TabInd + vbTab + "!" + TabInd + "����������;"
		elseif ObjParts(0) = "������������������" then
			AutoFillText = AutoFillText + TabInd + "���� " + CurrentLine + ".����������������() = 1 ����"
			AutoFillText = AutoFillText + TabInd + vbTab + "!" + TabInd + "����������;"
		end if

	Case "���� �� �������"
		AutoFillText = AutoFillText + TabInd + "���� " + CurrentLine + ".��������������() = 1 ����"
		AutoFillText = AutoFillText + TabInd + vbTab + "!" + TabInd + "����������;"

	Case "���� �� ���������"
		AutoFillText = AutoFillText + TabInd + "���� " + CurrentLine + ".����������������() = 1 ����"
		AutoFillText = AutoFillText + TabInd + vbTab + "!" + TabInd + "����������;"

	Case "rmtrade"
		AutoFillText = AutoFillText + CurrentLine + ".RMTrade, ""/D!"
	end select
end function
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

Private Sub DoAutoFill(doc, AutoText)
	If AutoText = "" Then Exit Sub
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = Replace(AutoText, "!", "")
	If InStr(AutoText, "!") <= 0 Then Exit Sub

	AddedStrings = Split(AutoText, vbCrLf)
	For i = 0 To UBound(AddedStrings) Step 1
		pos = InStr(AddedStrings(i), "!")
		If pos > 0 Then
			Col = pos - 1
			If i = 0 Then Col = Col + doc.SelStartCol
			doc.MoveCaret doc.SelStartLine + i, Col
			Exit Sub
		End If
	Next
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' kuntashov@yandex.ru, 2005/06/04
' ��������� ������ ������� Methods, �������� � ��� ������
' ������, ������������ � LeftPart
Private Function FilterMethodsList(Methods, LeftPart)
	' XXX �������� � ���, ���������� ����������� ����������
	' � ���������, � ������� �������� �������� ������, �����
	' �� ��������������� ������ �� ���������� �������������
	' � ������ � ������� ������ ���
	Dim Ret, MethArr
	Ret = ""
	MethArr = Split(Methods, vbCrLf)
	For i=0 To UBound(MethArr) Step 1
		Pos = InStr(1,MethArr(i),"#")
		If Pos <> 0 Then
			If UCase(Mid(MethArr(i),Pos+1,Len(LeftPart))) = UCase(LeftPart) Then
				Ret = Ret + MethArr(i) + vbCrLf
			End If
		End If
	Next
	FilterMethodsList = Ret
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'����������: 1 - �� ���������,
'            0 - �� ������� ��������� ������ �������,
'           -1 - ��� ������ ������ ������������ ����� ESC
Function MethodsList()
ProfilerEnter("MethodsList")
	MethodsList = 0
	doc = ""
	if Not CheckWindow(doc) then Exit Function

	Set SelObj = CreateObject("SelectValue.SelectVal")
	Set wshShell = CreateObject("wScript.Shell")

	DLLKIND = GetIniParam("CHOICELIST_KIND")

	ActiveWnd_HWnd = SelObj.GetActiveWnd()
	ModuleWnd_HWnd = Windows.ActiveWnd.HWnd

ProfilerEnter("PrepareModuleText")
	RSide = PrepareModuleText(doc.SelStartLine, doc.SelStartCol)
ProfilerLeave("PrepareModuleText")

	CurDocLine = doc.SelStartLine
	CurDocCol = doc.SelStartCol

	DlgType = 0

	Debug "CurrentLine", CurrentLine
	Debug "RSide", RSide

	if CurrentLine = "" then
		VsePredopredelennye
ProfilerLeave("MethodsList")
		exit Function
	elseif right(CurrentLine, 1) = "(" or right(CurrentLine, 1) = "," then
		SymbBefore = """"
		SymbAfter = """"
		Methods = GetListParamValue(lcase(CurrentLine))
	elseif right(CurrentLine, 1) = "=" then
		SymbBefore = """"
		SymbAfter = """"
		Methods = GetEqualValues(CurrentLine)
	else
ProfilerEnter("MethodsList 1")
		Set RE = New RegExp
		RE.IgnoreCase = True
		' �.�. ������ ����
		'	�������������("
		'	���� ��������������(���) = "
		'	(���� = ��� <> ����������)
		'	.��������������������("
		'
		RE.Pattern = "^.*((�������������|CreateObject|����������������������|GetEmptyValue|\.��������������������|\.���������������������|\.��������������������������)\s*\(\s*)|((����|If|����|While)\s+.*(��������������|ValueTypeStr).*(=|<>)\s*)""$"
		Set Matches = RE.Execute(CurrentLine)
		Set Matches2 = RE.Execute("."+RSide) ' artbear ����� ������ ����������� ���� �.��������������������("
ProfilerLeave("MethodsList 1")
ProfilerEnter("MethodsList 2")
		if (Matches.Count > 0) or (Matches2.Count > 0) then
			SymbBefore = ""
			SymbAfter = ""
			Methods = GetAllPossibleTypesList()

			ToAddDot = false ' ����� ����� �� �����

		else
			Debug "������ �� ������������� ������� ��� ������ �����", CurrentLine
ProfilerLeave("MethodsList 2")
ProfilerEnter("MethodsList 3")

			RE.Pattern = "//\s*:\s*"
			Set Matches = RE.Execute(CurrentLine)
ProfilerLeave("MethodsList 3")
ProfilerEnter("MethodsList 4")
			if (Matches.Count > 0) then
				SymbBefore = ""
				SymbAfter = ""
				Methods = GetAllPossibleTypesList()

				ToAddDot = false ' ����� ����� �� �����

			else
				Debug "������ �� ������������� ������� ��� ������ �����", CurrentLine
ProfilerLeave("MethodsList 4")
ProfilerEnter("MethodsList 5")

				Methods  = GetMethodsList(lcase(CurrentLine))
				if RSide <> "" Then ' kuntashov@yandex.ru, 2005/06/04
					Methods = FilterMethodsList(Methods, RSide)
				end if
ProfilerLeave("MethodsList 5")
			end if

		end if
	end if

	if Methods = "" then
	ProfilerLeave("MethodsList")
	 exit Function
	end if

	if DlgType = 0 then

		AttrName = SelectFromList(Methods, RSide)

	else
		AttrName = SelObj.SelectValue(Methods, "�������� �������", 1, vbCrLf, Windows.ActiveWnd.HWnd)
		AttrName = Replace(AttrName, vbCrLf, ",")
	end if

	if AttrName = "" then
		Debug "SymbBefore", SymbBefore
'		Debug "AttrName", AttrName
		Debug "SymbAfter", SymbAfter

		EscParam = GetIniParam("POINT_AFTER_ESC")
		if UCase(EscParam) = "��" and ToAddDot then
			Doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = "."
			Doc.MoveCaret doc.SelStartLine, doc.SelStartCol + 1
		end if
		MethodsList = -1 ' ������ ESC
ProfilerLeave("MethodsList")
		Exit Function

	elseif AttrName = "." then
		SymbBefore = ""
	else
		'AttrName = Replace(AttrName, "f)", "")
	end if

	if SymbBefore = "." and  not ToAddDot Then
		SymbBefore = ""
	end if

	ExtAction = ""
	if Instr(AttrName, "+") > 0 then
		ExtAction = Trim(RightFrom(AttrName, "+"))
		AttrName = Trim(LeftFrom(AttrName, "+"))
		CurrentLine = GetExprParts(CurrentLine + tDot)
		TabInd = vbCrLf + String(doc.SelStartCol - len(CurrentLine)-1, vbTab)
	end if

	if (Right(CurrentLine,1) = "." ) And (SymbBefore = ".") Then
		SymbBefore = ""
	end if
	if InStr(1, AttrName, "(f)") Then
		AttrName = Replace(AttrName, "(f", "(!")
		if Right(AttrName,1) = ")" Then
			AttrName = Replace(AttrName, ")", ");")
		end if
	end if

	Pos = InStr(AttrName,"!")
	if InStr(AttrName,"!") = 0 Then
		Pos = len(AttrName)
	else
		Pos = Pos - 1
	end if
	Pos = Pos + Len(SymbBefore)
	AttrName = Replace(AttrName, "!", "")

	Debug "SymbBefore", SymbBefore
	Debug "AttrName", AttrName
	Debug "SymbAfter", SymbAfter

	AttrName = SymbBefore + AttrName + SymbAfter

	ssc = doc.SelStartCol
	if RSide <> "" then ' kuntashov@yandex.ru, 2005/06/04
		ssc = ssc - Len(RSide)
	end if

	doc.range(doc.SelStartLine, ssc, doc.SelEndLine, doc.SelEndCol) = AttrName
	doc.MoveCaret doc.SelStartLine, ssc + Pos

	if ExtAction <> "" then
		AutoText = AutoFillText(ExtAction, CurrentLine, TabInd)
		DoAutoFill doc, AutoText
	end if

	MethodsList = 1

	if DLLKIND <> "2" And DLLKIND <> "3" then
		if ActiveWnd_HWnd <> SelObj.GetActiveWnd() then
			SelObj.ActivateWindow ActiveWnd_HWnd
			'ShowTooltip 0
			SelObj.ShowCurs ModuleWnd_HWnd
		end if
	end if
	ShowTooltip 0
ProfilerLeave("MethodsList")
End Function

' ���������� ��������� �������� �� ����������
Sub ShowTooltip(param)
	if UCase(GetIniParam("SHOWTOOLTIPS")) = "��" then
		Set wshShell = CreateObject("wScript.Shell")
		'wshShell.sendKeys "{ESC}" '����� �������� ESC ��� �� �����, �� �����-�� ���� �����
		wshShell.sendKeys "^+ " '���������� ��������� �������� �� ����������
  	end if
End Sub

'
' �������� ��� ������ � ints-������ ��� ������� 1�++
'
Sub Update1CppInts()
	On Error Resume Next
	Set CPPAnalyser = CreateObject("OpenConf._1CPP")
	On Error Goto 0
	If CPPAnalyser Is Nothing Then
		Message "�� ���������� ��� �� ��������������� ��������� SyntaxAnalysis.wsc", mRedErr
		Exit Sub
	End If

	' ����-����� ����� �������� � ����������� �����, ����� �� ���������� ����� �����
	if not fso.FolderExists(Get1CppIntsFolder()) then
		set f = fso.CreateFolder(Get1CppIntsFolder())
	end if

	CPPAnalyser.SetConfig(Configurator)
	CPPAnalyser.UpdateClassesInfo()
	Dim i
	For i = 0 to CPPAnalyser.ClassCount - 1
		Set Info = CPPAnalyser.ClassInfo(i)
		FileName = Get1CppIntsFolder & "\" & Info.Name & ".ints"
		Status FileName

		Set IntsFile = fso.CreateTextFile(FileName)
		Set MethodList = CreateObject("Scripting.Dictionary")
		AddCppMethods Info, IntsFile, MethodList
		IntsFile.Close

		' ������ ����-���� �� ����� Intell ������, ����� �� ���� ������������
		FileName = DataDir & "\" & Info.Name & ".ints"
		Debug "������ ������ ����",  FileName
		if fso.FileExists(FileName) then
			fso.DeleteFile FileName, true
		end if
	Next
End Sub

Private Sub AddCppMethods(ClassInfo, IntsFile, MethodList)
	If CPPAnalyser Is Nothing Then Exit Sub
	Debug "����� ", ClassInfo.Name

	Dim i
	For i = 0 to ClassInfo.BaseClasses.Size()-1 Step 1
		BaseName = ClassInfo.BaseClasses.Item(i)
		Set BaseClass = Nothing
		On Error Resume Next
		Set BaseClass = CPPAnalyser.ClassInfo(BaseName)
		On Error Goto 0
		If not BaseClass Is Nothing Then AddCppMethods BaseClass, IntsFile, MethodList

		' ��� �������� �������� ������ ������ ������� ��� ��� ������ � ���� ������ 1�++
		If BaseClass Is Nothing Then AddStandartBaseClassTo1CppInts ClassInfo.Name, IntsFile, BaseName, MethodList

	Next

	If ClassInfo.Methods.Size() = 0 Then ClassInfo.FillMethodList
	'add exported vars
	For i = 0 to ClassInfo.Vars.Size() - 1
		VarName = ClassInfo.Vars.Item(i)
		If not MethodList.Exists("var " & VarName) Then
			MethodList.Add "var " & VarName, "var " & VarName
			IntsFile.WriteLine "0000 " & VarName
		End If
	Next
	'add exported methods
	For i = 0 to ClassInfo.Methods.Size() - 1
		method = ClassInfo.Methods.Item(i)
		func = ""
		If Left(method, 1) = "f" Then func = "f"
		method = Mid(method, 3)
		If not MethodList.Exists(method) Then
			MethodList.Add method, method
			IntsFile.WriteLine "0000 " & method & "(" & func & ")"
		End If
	Next
End Sub

' ��� �������� �������� ������ ������ ������� ��� ��� ������ � ���� ������ 1�++
'
' ����� ����������� ������ �� ���� ������ ���������� ��� ��������, ���� ����� ����������, ��������, ����������.���������� ��� ��������.��������������
' ����� ����������� �������� ����� �����������
'
Private Sub AddStandartBaseClassTo1CppInts(Class1CppName, IntsFile1Cpp, ClassName, MethodList)
	Debug "������� � ����� ", ClassName

	FileName = GetIntsFilePath(ClassName)
	if FileName <> "" then
		CopyIntsFileToDestFile FileName, IntsFile1Cpp, MethodList
		Exit Sub
	end if

	' �������� �������� ������ � �������� �����������, ��������, ����������.����������

	ClassParts = Split(ClassName, ".")
	if UBound(ClassParts) = 0 then
			CommonScripts.Error "��� �������� ������ " & ClassName & " (��� ������� " & Class1CppName & ") ��� ���������������� ints-�����!"
			Exit Sub
	end if

	MainType = lcase(ClassParts(0)) '  ��������, ��� ����������.���������� ��� ����������

	Debug "������� � ����� ������", MainType

	' ��� ���������� ?
	if not MethodsBehaviour.Exists(MainType) then
			CommonScripts.Error "������� ����� " & ClassParts(0) & " (��� ������� " & Class1CppName & ") �� �������� ����������!"
			Exit Sub
	end if

	' ������� ������ ����� ������ � ��������� ��� ���� ������, ��������, ��� ���������� ��� ��������
	Debug "������ ����� ������ ���� ������", MainType

	FileName = GetIntsFilePath(MainType)
	if FileName = "" then
			CommonScripts.Error "��� �������� ������ " & ClassParts(0) & " (��� ������� " & Class1CppName & ") ��� ���������������� ints-�����!"
			Exit Sub
	end if
	CopyIntsFileToDestFile FileName, IntsFile1Cpp, MethodList

	' � ������ ������� �������� ����������� �����������, ��������, ����������.����������
	Debug "������ �������� ����������� �����������", ClassName

	Dim BhvInfoParts,  Attrs

	BhvInfoParts = Split(MethodsBehaviour.Item(MainType), "|")

	Attrs = ""
	Select Case BhvInfoParts(0)
	Case "ATTR":
		Attrs = GetAttrs(CLng(BhvInfoParts(1)), LCase(ClassParts(1)), 0, CLng(BhvInfoParts(2)))

'TODO �������� �� ������ ���� ���������� ������� ??

	End Select

	if Attrs = "" then
			CommonScripts.Error "��� �������� ������ " & ClassName & " (��� ������� " & Class1CppName & ") �� ������� �������� �������� �����������! ��������, ������ ����������� ��� � ������� !"
			Exit Sub
	end if

	AttrsParts = Split(Attrs, vbCRLF)
	Dim i, i2
	For i = 0 to UBound(AttrsParts)

		VarName = AttrsParts(i)
		' ������ ����������� �����
		VarName = Mid(VarName, 4)

		If not MethodList.Exists("var " & VarName) Then ' ������� ������������
			MethodList.Add "var " & VarName, "var " & VarName
			IntsFile1Cpp.WriteLine "0000 " & VarName
		End If
	Next

End Sub ' AddStandartBaseClassTo1CppInts

' �������� ����������� �� ������ ����-����� � ������
' ����� ��������� ������������,  ��������� ������� MethodsList
Sub CopyIntsFileToDestFile(IntsFileName, DestFile, MethodList)
  Dim i

	Set IntsFile = fso.OpenTextFile(IntsFileName)
	Str = IntsFile.ReadAll

	IntsFile.Close
	IntsFile = 0

	StrParts = Split(Str, vbCrLf)
	for i = 0 to UBound(StrParts)
		item = Trim(StrParts(i))
		if item <> "" then
			item = Mid(item,6) ' ������� �������� �����

			If not MethodList.Exists("var " & item) Then ' ������� ������������
				MethodList.Add "var " & item, "var " & item
				DestFile.WriteLine "0000 " & item
			End If
		End If

	next

End Sub ' CopyIntsFileToDestFile

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ������� ��������. ���������� ����� ������� ������ "����� �����"
' �������� ����� ������/������� ����� ������.
' ���������� ����� ������ ����, ����������� vbCr ��� vbLF ��� vbCrLf ��� vbTab
function Telepat_OnShowMemberList(Line, Col)
	Telepat_OnShowMemberList=""
	doc = ""
	if Not CheckWindow(doc) then Exit Function

	MaxLineCount = GetIniParam("MAX_LINES_CALL_INTELL")
	if MaxLineCount = "default" then
		MaxLineCount = 9999
	else
		MaxLineCount = CLng(MaxLineCount)
	end if

	' ��� ������� ������� �������� Intellisence ������ �������������
	if doc.LineCount > MaxLineCount then Exit Function

	MethodsList

'	PrepareModuleText Line, Col
'
'	CurDocLine = Line
'	CurDocCol = Col
'
'
'	Methods  = GetMethodsList(lcase(CurrentLine))
'	params = split(Methods, vbCrLf)
'	for i = 0 to ubound(params)
'		dd = instr(params(i), "#")
'		if dd > 0 then params(i) = mid(params(i), dd + 1)
'	next
'	Telepat_OnShowMemberList = Join(params, vbCr)
End Function

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' ������� ���������� ����� �������� ������ �� ������.
' ��������� ��������������� �����.
' �����, ����������� � ���������� �������, ���������� � InsertType=15 (������)
Sub Telepat_OnInsert(InsertType, InsertName, Text)
	Select Case InsertType
		Case 10         ' ����� 1�
			' ����� ��� �� �������� ��������� ����������� �������������...
			'	� ����� ������ ������ ���� ��������, � �.�. � ������� 1�++

			If InsertName = "�������������" Then Text = "�������������(""!"");"
			' ����� ������ ����� ����� ���������� � ������, �������� ������ ����� ��������� �������
			gTimerIns = CfgTimer.SetTimer(1, False)

		Case 11         ' �������� �����
            If InStr(",���,��,", InsertName) > 0 Then Text = UCase(Text)

        Case 14         ' ���������������� �����
        	If InsertName="���������_��������������" or InsertName="�������_������������" Then
        		ProcName=InputBox("������� ��� ������","�����", "��")
        		Text=Replace(Text,"!(",ProcName & "(!") & vbTab & "// " & ProcName
        	End If

		Case 15		 ' ����������� � ������
			if Instr(text,"f)")>0 or Instr(text,"(?)")>0 then SendCommand 22502	' ������� ���������
			Text = Replace(Text, "f)", "")
			Text = Replace(Text, "?)", "")
			ExtAction = ""
			if Instr(Text, "+") > 0 then
				ExtAction = Trim(RightFrom(Text, "+"))
				Text = Trim(LeftFrom(Text, "+"))
				CurrentLine = GetExprParts(CurrentLine + ".")
				TabInd = vbCrLf + String(CurDocCol-len(CurrentLine), vbTab)
				if ExtAction <> "" then
					text = text + AutoFillText(ExtAction, CurrentLine, TabInd)
				end if
			end if
	End Select
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
'	���������� ��� ������������� � ����������� ������� ��������
'	����� ������ ����� ����� ���������� � ������, �������� ������ ����� ��������� �������
'
Sub Configurator_OnTimer(timerID)
    If timerID = gTimerIns Then
		doc = ""
		if CheckWindow(doc) then
			CfgTimer.KillTimer gTimerIns		'������� ������
			MethodsList
		end if
    End If
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
Sub InitTelepat(EmptyParam)
	fname = FSO.GetParentFolderName(Scripts.Path(SelfScript.Name)) & "\RunAllIntellisence.vbs"

	'If not FSO.FileExists(fname) and UCase(GetIniParam("TELEPAT")) = "��" Then
	If UCase(GetIniParam("TELEPAT")) = "��" Then
		Set t = Plugins("�������")  ' �������� ������
		If Not t Is Nothing Then	' ���� "�������" ��������
			' ����������� ������ � �������� �������
			SelfScript.AddNamedItem "Telepat", t, False
		end if
	End If
End Sub

'
' ��������� ������������� �������
'
Private Sub Init()
	InitTelepat 0

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

	'' ��������: ��������� 2 ������ ����������� � ����� �������, ����� �� ���������� !
	'
	'' ��� �������� ������� �������������� ���
	'c.AddPluginToScript SelfScript, "�������", "Telepat", Telepat

	SelfScript.AddNamedItem "CommonScripts", c, False

	' ��������� �������
	CommonScripts.SetQuietMode not DebugFlag

	' ����� ������� �������....
	ReloadDictionary ' trdm
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
'
'	�������� ��� ������ ������ ���� ��������� ����� + ������ 1�++
'	��������� � ����� ����� ������������
'
Sub SelectAndPasteTypeFromPossibleTypes()

	doc = ""
	if Not CheckWindow(doc) then Exit Sub

	strType = SelectTypeFromAllPossibleTypesList()

	if strType <> "" then
		Doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = strType
		Doc.MoveCaret doc.SelStartLine, doc.SelStartCol + Len(strType)
	end if

End Sub ' SelectAndPasteTypeFromPossibleTypes

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' artbear@bashnet.ru, 2005/06/15
'
'	�������� ����������� ���� ���������� ��� ����� ������� ������������� ��������
'		1) ������������ ����� �� ������ ��������� ����� � ������� 1�=1�+1;
'		2) ��������� � ����� ������ " // : �������"
'
Sub AddTypeDefinition()

	doc = ""
	if Not CheckWindow(doc) then Exit Sub

	strType = " // : "
	strType = strType + SelectTypeFromAllPossibleTypesList()

	if strType <> "" then
		Doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = strType
		Doc.MoveCaret doc.SelStartLine, doc.SelStartCol + Len(strType)
	end if

End Sub ' AddTypeDefinition

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' trdm@mail333.com, 2005/10/10
' ��������� ��� ������� �� ��������� �����
'
Private Sub ReloadOneDict(Dict, FileName)

	if not FSO.FileExists(FileName) then Exit Sub

	Set FileStream = fso.OpenTextFile(FileName, 1, false)
	if FileStream.AtEndOfStream = true then  Exit Sub

	AllText = FileStream.ReadAll()
	AllDictionatyItem = Split(AllText, vbCrLf)
	Dividor = ","
	for i = 0 to UBound(AllDictionatyItem)
		If AllDictionatyItem(i)<>"" Then
			Parts = Split(AllDictionatyItem(i),Dividor)
			Status "" + Parts(0)

			if UBound(Parts) = 1 Then
				if Not Dict.Exists( Lcase( Parts(0))) then
					Dict.Add Lcase(Parts(0)), Parts(1)
				end if
			end if
		end if
	next
End Sub

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
' trdm@mail333.com, 2005/10/10
'
' ��������� �������, "OtherTypesDefine" � "TypesStructureExt" �� ��������� ������
Sub ReloadDictionary()
	FileName = DataDir + "\TypesStructureExt.txt"
	ReloadOneDict TypesStructureExt, FileName

	FileName = DataDir + "\OtherTypesDefine.txt"
	ReloadOneDict OtherTypesDefine, FileName
End Sub ' ReloadDictionary

' ���������� ���������
Sub Debug(ByVal title, ByVal msg)
	CommonScripts.Debug title, "<" & msg & ">"
End Sub'Debug

' �������� ����� ���������� ���������
Sub EnableDebugMessages()
	DebugFlag = true
	CommonScripts.SetQuietMode not DebugFlag
End Sub ' EnableDebugMessages

' ��������� ����� ���������� ���������
Sub DisableDebugMessages()
	DebugFlag = false
	CommonScripts.SetQuietMode not DebugFlag
End Sub ' DisableDebugMessages

' ������ � ����������� ��� ����������� �������
'Dim Profiler
Private Sub ProfilerEnter(name)
'	Set Profiler = CreateObject("SHPCE.Profiler")
'	Profiler.StartPiece(name)
End Sub

Private Sub ProfilerLeave(name)
'	Set Profiler = CreateObject("SHPCE.Profiler")
'	Profiler.EndPiece(name)
'	if Name = "MethodsList" then
'		ProfilerShow
'	end if
End Sub

Private Sub ProfilerShow()
'	Set Profiler = CreateObject("SHPCE.Profiler")
'	Set Results = Profiler.Results
'	For Each R In Results
'	'	Message R & "/" & TypeName(R) & "/" & IsObject(R) & "/" & VarType(R)
'		Message R
'	Next
End Sub

Init  ' ������������� �������
