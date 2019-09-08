'--------------------------------------------------------------
'    ����� �������� ��� ��������� ����� ��������� ������������
'--------------------------------------------------------------
'
'������: $Revision: 1.43 $ 
'
'    GoToObject:
' ������, ������� ��������� ������ (����������, ��������, ������, �����, ���������, ������ ������ 1�++), 
' ��� �������� ��������� ��� ��������.
' ��� ����� ���� ������ (����������.�����������) ��� ��� ���� (�����������)
' ��� ������������ � �������� ����� ������� ����� ������ ��� ��������. ��� ���������� ����� ������� 
' ����� ��� ������.
'
' ���� ��� �������� ��������� ��������� ���� <����������>.<������|���>, �� �������� ������� ���������� 
' ��� ������� (��� ������ Intellisence.vbs), � ���� ��� �����, �� ����������� codejump � ������ ������,
' � ����� ������.
'
' ���� ��� �������� ��������� ��������� ���� '���.�����()' �� �������� ������� ���������� ����� �������������
' ���������� '���' � ������� � ���� ������
'
' ���� ��� �������� ��������� ������ ������, ��� ���������, ��� �������� �� ������� ����������, �� ��������
' ������� ������� � ������ ����������, �� ������� ������������� ������.
'
' ���� � ������������ ������ �� �������, �� ������ �������� ������� ������� ����� (��� ����). ������� �����
' ������ � �������� ��, � �������� ExtForms, � ��������� �� ����������������� ����� �� ������� ExtFormsTree 
' �� ������ �������, � ����� ���������� �� ���� ������������ �������� ��.
' ���� � ��� �� �������, �� ����������� ������ "������� ���� �� ��������� ����������������" �� AlexQC (���� �� ������� ����)
'
' ���� ������� ���� �� ���������, �� �������� ������� ��������� ������ Scripts("NavigationTools").GoToFormula - 
' ������� � ���������/�������, ���������� �� ������� �������.
'
' ��� ������ �������� �� ������ ������������ ActiveX ������ SvcSvc.Service, ������� ����� ����� �� ������ �������� �� artbear.
'--------------------------------------------------------------
'   ������� ��� ��������� �� ��������� (������, ������, ������� ������) ������� �����
'    ToggleWorkbookPage:  ������ ������ ��� ������������ ������� ������/������, ��� � Delphi �� F12
'    GoToDialogPage:      ������������ �� �������� �������
'    GoToModulePage:      ������������ �� �������� ������
'    GoToMXLPage:         ������� � ������� �������. ���� �������� ���������, �� ������� ������ ��� ������
'    GoToPage:            ������� ������ ���� �������, � ������������ ������� � ��������� ��������.
'
' �����: ������� ����� aka ADirks
' e-mail: <adirks@ngs.ru>
'
'===========================================================================================================================
' ������� OpenFileFromClipboard, GetFileNameFromClipboard � Configurator_OnFileDialog
' ������ ��� �������� ������, ����� ��� ����� ����� � clipboard
'
' � ������� GetFileNameFromClipboard() ������������ ������ WshExtra.Clipboard. COM-������ � ���� ��������
' ����� ����� �� ������ http://www.winnetmag.com/Files/07/23601/23601.zip
'
' �����: ����� ������� aka artbear
' e-mail: artbear@bashnet.ru
' ICQ: 265666057

Dim NameDelimiters
Dim PathDelimiters
Dim RecursiveSearchDepth
Dim fso
Dim SA
Dim flSelSensitive

'======= ������������� =============================

' ��������� ������������� �������
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

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set SA = Nothing
	On Error Resume Next
	Set SA = CreateObject("OpenConf._1CPP")
	SA.SetConfig(Configurator)
	On Error Goto 0

	NameDelimiters = vbTab & " ,;:|#=+-*/%?<>\()[]{}!~@$^&'"""
	PathDelimiters = """"
	RecursiveSearchDepth = 3
	flSelSensitive = true
End Sub
 
Init ' ��� �������� ������� ��������� �������������

'===================================================


Sub GoToObject()
	If IsTextWindow Then
		If TryOpenIncludeFile() Then Exit Sub
		If TryClassMethodJump Then Exit Sub
		If TryToOpenObject("") Then Exit Sub
		If TryVarDefJump() Then Exit Sub
		If TryToOpenFile("") Then Exit Sub
		If TryGotoTree() Then Exit Sub

	Else
		If TryGoToFormula Then Exit Sub
		If TryGotoTree Then Exit Sub
		If TryGoToSelectedMDTreeItemType Then Exit Sub		
	End If

	'���� � ��� ������ ��� � �� ���������, �� �������� ���������� ���������, � ���� ��� ��� �����, �� ������� ���
	OpenFileFromClipboard
End Sub 'GoToObject

Private Function TryGoToSelectedMDTreeItemType()
	TryGoToSelectedMDTreeItemType = true 
	On Error Resume Next
	TryGoToSelectedMDTreeItemType = Scripts("NavigationTools").GoToSelectedMDTreeItemType()
	On Error Goto 0
End Function	

Private Function TryOpenIncludeFile()
	TryOpenIncludeFile = false
	On Error Resume Next
	Set PrevWin = Windows.ActiveWnd
	Scripts("������� ���� �� ��������� ����������������").OpenIncludeFile
	If PrevWin <> Windows.ActiveWnd Then TryOpenIncludeFile = true
	On Error Goto 0
End Function

Private Function TryGotoFormula()
	TryGotoFormula = false
	On Error Resume Next
	TryGotoFormula = Scripts("NavigationTools").GoToFormula
	On Error Goto 0
End Function

Sub ClassMethodJump
	TryClassMethodJump
End Sub

Private Function TryClassMethodJump
	TryClassMethodJump = false
	
	If SA Is Nothing Then Exit Function

	Set Doc = CommonScripts.GetTextDocIfOpened(false, true)
	If Doc Is Nothing Then Exit Function

	Line = Doc.SelStartLine
	Col = Doc.SelStartCol

	'������� ��������� �� �����
	Str = Doc.Range(Line)
	pos = InStrRev(Str, ".", Col+1)
	If pos <= 0 Then Exit Function
	LeftPart = Left(Str, pos-1)
	RightPart = Mid(Str, pos+1)
	Expr = GetObjectName(LeftPart, Len(LeftPart), NameDelimiters)
	MethodName = GetObjectName(RightPart, 0, NameDelimiters)
	'Message Expr & ":" & MethodName

	If Expr = "" or MethodName = "" Then Exit Function

	'���������� ���������� ��� ����� �����
	ClassName = ""
	On Error Resume Next '���� ������� ������� ��� ������� ���, �� � ClassName ��������� ������ ������, � �� ���������� ����-�����
	Set Intellisence = Scripts("Intellisence")
	ClassName = Scripts("Intellisence").GetExprType(Doc, Expr)
	On Error Goto 0

	If ClassName = "" Then Exit Function
	'���� ��� �����, �� ����� ���������� � ��
	Set ClassInfo = Find1CPPClass(ClassName)
	'Message ClassName & ": " & ClassInfo
	If ClassInfo Is Nothing Then Exit Function

	MethodLine = FindClassMethod(ClassInfo, MethodName, ClassDoc)
	Doc.Open '������������ �������� ������, ����� � ����� ������� ����������� ���������� �������
	If MethodLine < 0 Then Exit Function

	'� �������-�� �������� ����������� ������ � ������
	CommonScripts.Jump MethodLine, -1, -1, -1, ClassInfo.Location
	If UCase(Right(ClassInfo.Location, 3)) = "@MD" Then
		ClassDoc.Open '������-�� jump � ��������� �� ������������ ����. ���������� ��� ��� ���.
	End If
	Windows.ActiveWnd.Caption = "����� " & ClassInfo.Name

	TryClassMethodJump = true
End Function

Function FindClassMethod(ClassInfo, MethodName, ClassDoc)
	FindClassMethod = -1
	Set ClassDoc = ClassInfo.OpenModule()
	If not ClassDoc Is Nothing Then
		FindClassMethod = CommonScripts.FindProc(ClassDoc, MethodName)
		If FindClassMethod >= 0 Then Exit Function
	End If

	Dim i
	For i = ClassInfo.BaseClasses.Size()-1 To 0 Step -1
		BaseName = ClassInfo.BaseClasses.Item(i)
		Set BaseClass = Nothing
		On Error Resume Next
		Set BaseClass = SA.ClassInfo(BaseName)
		On Error Goto 0

		If not BaseClass Is Nothing Then
			found = FindClassMethod(BaseClass, MethodName, ClassDoc)
			If found >= 0 Then
				FindClassMethod = found
				Set ClassInfo = BaseClass
				Exit Function
			End If
		End If
	Next
End Function

'������� � ������, ��� ���� ���������� ����������
Sub VarDefJump
	TryVarDefJump
End Sub

Private Function TryVarDefJump
	TryVarDefJump = false
	Set Doc = CommonScripts.GetTextDocIfOpened(false, true)
	If Doc Is Nothing Then Exit Function

	Line = Doc.SelStartLine
	Col = Doc.SelStartCol
	Expr = GetObjectName(Doc.Range(Line), Col, NameDelimiters)
	If Expr = "" Then Exit Function

	pos = InStr(Expr, ".")
	If pos > 0 Then Expr = Left(Expr, pos)

	'���������� ���������� ��� ����� �����
	strType = ""
	TypeDef_Line = -1
	On Error Resume Next '���� ������� ������� ��� ������� ���, �� � strType ��������� ������ ������, � �� ���������� ����-�����
	Set Intellisence = Scripts("Intellisence")
	strType = Scripts("Intellisence").GetExprType(Doc, Expr)
	TypeDef_Line = Intellisence.LineNumber(Intellisence.TypeDef_Pos)
	On Error Goto 0
	If TypeDef_Line < 0 Then Exit Function

	CommonScripts.Jump TypeDef_Line
	TryVarDefJump = true
End Function

Sub ToggleWorkbookPage()
	If Windows.ActiveWnd Is Nothing Then Exit Sub
	Set Doc = Windows.ActiveWnd.Document
	If Doc Is Nothing Then 
		' ����� ������ �� Ctrl+Tab, �� ��������� ����������� �������������
		' ����� �������� (��-WorkBook) ������ ����� � ������� Ctrl+Tab
		Configurator.CancelHotKey = True
		Exit Sub
	End If
	If Doc.Type <> docWorkBook Then 
		Configurator.CancelHotKey = True
		Exit Sub
	End If
	Set Page = Doc.Page(Doc.ActivePage)
	
	If Page.Type = docText Then
		GoToDialogPage
	ElseIf Page.Type = docDEdit Then
		'GoToModulePage
		Doc.ActivePage = "������"
	End If
End Sub

Sub GoToDialogPage()
	GoToPageByType Nothing, docDEdit
End Sub

Sub GoToModulePage()
	GoToPageByType Nothing, docText
End Sub

Sub GoToMXLPage()
	GoToPageByType Nothing, docTable
End Sub

Sub GoToPage()
	GoToPageByType Nothing, -1
End Sub

Sub GoToPageByType(WB, nPageType)
	If WB is Nothing Then
		If Windows.ActiveWnd Is Nothing Then Exit Sub
		Set WB = Windows.ActiveWnd.Document
	End If
	If WB is Nothing Then Exit Sub
	If WB.Type <> docWorkBook Then Exit Sub
	
	PageList = ""
	Dividor = ""
	For i = 0 To WB.CountPages - 1
		If nPageType < 0 OR WB.Page(i).Type = nPageType Then
			PageList = PageList & Dividor & WB.NamePage(i)
			Dividor = vbCrLf
		End If
	Next
	If PageList = "" Then Exit Sub

	If InStr(PageList, vbCrLf) = 0 Then
		PageName = PageList
	Else
		PageName = CommonScripts.SelectValue(PageList, "��������")
	End If
	If PageName = "" Then Exit Sub
	
	WB.ActivePage = PageName
End Sub

' �������� ������� ������ ��� ��������
' WhatToOpen - ��� ��������� ������. �.�. "�����" ��� "�����.������". �� ��������� - "�����.������"
' ���������� true, ���� ������ ������, false - ���� �� ������
Function TryToOpenObject(WhatToOpen)
	Dim ObjType, ObjName, NameToOpen

	TryToOpenObject = false

	If Trim(WhatToOpen) = "" Then WhatToOpen = "�����.������"

	Set Doc = CommonScripts.GetTextDocIfOpened(false, true)
	If Doc Is Nothing Then Exit Function

	'[*]MetaEditor, 20.06.2005
	'���� ���� ���������� �����, �� �� ������ ��� ������������ ������� ������� ����� ���
	'����������� � Init - flSelSensitive = false
	If (Doc.SelStartCol <> Doc.SelEndCol) and flSelSensitive Then
		FullName = doc.Range(doc.SelStartLine,doc.SelStartCol,doc.SelEndLine,doc.SelEndCol)
	Else           
		FullName = GetObjectName(Doc.Range(Doc.SelStartLine), Doc.SelStartCol, NameDelimiters)
	End If
	'_

	If TryToOpenClass(FullName) Then
		TryToOpenObject = true
		Exit Function
	End If

	If (Doc.SelStartCol <> Doc.SelEndCol) and flSelSensitive Then
		NameToOpen = FullName
	Else 
		Set Doc = Nothing
		On Error Resume Next
		Set Doc = Documents("" & FullName & "." & WhatToOpen)
		On Error Goto 0
		if Doc Is Nothing Then
			Str = FullName
			Do While GetIDFromString(Str, NameToOpen)
				NameToOpen = FindObject(NameToOpen)
				If NameToOpen <> "" Then Exit Do
			Loop
		Else
			NameToOpen = FullName
		End If
	End If
	
	'stop 
	'[*]MetaEditor, 21.06.2005 - ������� ��� �� ������, ���� �� ����
	if InStr(NameToOpen,",") = 1 then 
		NameToOpen = Mid(NameToOpen,2)
	end if
	if InStr(NameToOpen,",") > 0 then 
		NameToOpen = CommonScripts.SelectValue(NameToOpen,"������� ��������� �������� ��",",")
	end if
	'_
	
	'message NameToOpen
	If NameToOpen = "" Then Exit Function '�� ����� �������

	'������������� ������ ���������, ���� ���������
	ListForms = WhatToOpen
	Str = NameToOpen
	If GetIDFromString(Str, ObjType) Then
		If ObjType = "����������" or ObjType = "������" Then
			GetIDFromString Str, ObjName
			ListForms = GetListForms(ObjType, ObjName, WhatToOpen)
			If ObjType = "����������" Then
				ListForms = WhatToOpen & "," & Replace(WhatToOpen, "�����.", "����� ������.") & "," & ListForms
			ElseIf ObjType = "������" Then
				ListForms = ListForms
			Else
				ListForms = WhatToOpen & "," & ListForms
			End If
		ElseIf ObjType = "��������" Then
			ListForms = ListForms & "," & "������ ���������"
		ElseIf ObjType = "���������"  or ObjType = "�������" or ObjType = "�����������" or ObjType = "������������" Then	
			ListForms = "" '������ ������� � ������ ��
		End If
	End If  
	
	If ObjType <> "" and ObjName <> "" Then
		ListForms = ObjType & "." & ObjName & "," & ListForms
	ElseIf InStr(NameToOpen, ".") > 0 Then
		ListForms = NameToOpen & "," & ListForms
	End If
	
	If InStr(ListForms, ",") > 0 Then
		ListForms = Replace(ListForms, ",", vbCrLf)
		WhatToOpen = CommonScripts.SelectValue(ListForms, "��� ����� ���������?", vbCrLf, true)
		If WhatToOpen = "" Then
			TryToOpenObject = true
			Exit Function
		End If
	Else
		WhatToOpen = ListForms
	End If
	If WhatToOpen = "" Then
		Exit Function
	End If

	If WhatToOpen = FullName or WhatToOpen = NameToOpen Then
		Set NavigationTools = scripts("NavigationTools")
		NavigationTools.GoToMDTreeItem WhatToOpen, 0, true
		TryToOpenObject = true
	Else
		nLine = 0
		ModuleName = NameToOpen & "." & WhatToOpen
		Set Doc = Nothing
		On Error Resume Next
		Set Doc = Documents(ModuleName)
		On Error Goto 0
		If Not Doc Is Nothing Then
			If Doc.IsOpen Then
				Doc.Open
				TryToOpenObject = true
				Exit Function
			End If
			nLine = CommonScripts.FindProc(Doc, "�����������", "���������")
		End If
		If nLine < 0 Then nLine = 0
		TryToOpenObject = CommonScripts.Jump(nLine, -1, -1, -1, ModuleName)
	End If
End Function  'TryToOpenObject

'����������� ��� ������� �� �����
Function GetIDFromString(Str, ID)
	GetIDFromString = true
	Do
		pos = InStr(Str, ".")
		If pos = 0 Then
			ID = Trim(Str)
			Str = ""
			If Len(ID) = 0 Then
				GetIDFromString = false
				Exit Function
			End If
		Else
			ID = Left(Str, pos-1)
			Str = Mid(Str, pos+1)
		End If
	Loop while ID = ""
End Function

' ���������� ������ ��� ������� ��� ��������. � �������� ����� ������ ��������, ���������� ��������� �������, �
' ������������ ��������� � Delimiters
Function GetObjectName(Line, CursorPos, Delimiters)
	Dim Col1, Col2, TextLen

	GetObjectName = ""

	TextLen = Len(Line)
	Col1 = CursorPos
	do while Col1 > 0
		If InStr(Delimiters, Mid(Line, Col1, 1)) > 0 Then
			Col1 = Col1 + 1
			Exit Do
		End If
		Col1 = Col1 - 1
	loop
	If Col1 = 0 Then Col1 = 1
	Col2 = Col1
	do while Col2 <= TextLen
		If InStr(Delimiters, Mid(Line, Col2, 1)) > 0 Then
			Col2 = Col2 - 1
			Exit Do
		End If
		Col2 = Col2 + 1
	loop
	If Col2 > TextLen Then Col2 = TextLen

	'Message "" & Col1 & ":" & Col2, mNone
	GetObjectName = Mid(Line, Col1, Col2 - Col1 + 1)
End Function 'GetObjectName

' ���� ������ �� �����, ����� �� �������� ���. �.�., ���� �������� ������ ������������� (������, "�����������").
' ���� ������ ������, �� ������������ ������ ��� ("����������.�����������"), ���� ��� - �� ������ ������
' � � ObjType ������������ ��� �������
Function FindObject(ByVal ObjName)
	FindObject = "" 
	Set TypeObjs = Metadata.TaskDef.Childs
	For i = 0 To TypeObjs.Count - 1
		Set Objs = TypeObjs(i)
		For j = 0 To Objs.Count - 1
			Set Obj = Objs(j)
			If Obj.IsValid Then
				'Message TypeObjs.Name(i, true) & "." & Obj.Name, mNone
				If UCase(Obj.Name) = UCase(ObjName) Then
					'[*]MetaEditor, 21.06.2005 - �������� ������ ��� � ����������� ������� 
					'�������� ����� � ������������ ���� ���������� ���������� ���� ��������.������ � �������.������
					'� ����� � ������ �� ��������, �� ��������� �������� ������� ��������
					FindObject = FindObject & "," & TypeObjs.Name(i, true) & "." & Obj.Name
					'Exit Function
				End If
			End If
		Next
	Next
End Function 'FindObject

Function GetListForms(ObjType, ObjName, ByVal WhatToOpen)
	GetListForms = ""
	comma = ""
	Prefix = "�����������"
	If ObjType = "������" Then
		Prefix = "�����"
	End If
	If InStr(WhatToOpen, "������") > 0 Then
		WhatToOpen = ".������"
	Else
		WhatToOpen = ".������"
	End If

	Set Forms = Metadata.TaskDef.Childs("" & ObjType)("" & ObjName).Childs("�����������")
	For i = 0 To Forms.Count - 1
		Set Form = Forms(i)
		If Form.IsValid Then
			GetListForms = GetListForms & comma & Prefix & "." & Form.Name & WhatToOpen
			comma = ","
		End If
	Next
End Function 'GetListForms

' �������� ������� ������� ����� ��� ��������
' WhatToOpen - ������������
Function TryToOpenFile(WhatToOpen)
	TryToOpenFile = false

	Set Doc = CommonScripts.GetTextDocIfOpened(false, true)
	If Doc Is Nothing Then Exit Function

	FullName = GetObjectName(Doc.Range(Doc.SelStartLine), Doc.SelStartCol, PathDelimiters)
	Name = fso.GetFileName(FullName)
	If Len(Trim(Name)) = 0 Then Exit Function ' ���� ����� ����� ���, ����� ������ ? 
 	' ������, ��� ������� � ����� � ������� � ����� ����� ���� �� �����
 	If InStr(Name, ";") > 0 Then Exit Function
 	If InStr(Name, ",") > 0 Then Exit Function

	'���� ���������� �� ������, �� ��������� .ert
	If fso.GetExtensionName(Name) = "" Then
		Name = Name & ".ert"
		FullName = FullName & ".ert"
	End If

	'���������� ����� ���� �� ������������� ����������
	FileToOpen = FullName
	If Not fso.FileExists(FileToOpen) Then FileToOpen = IBDir & FullName
	If Not fso.FileExists(FileToOpen) Then FileToOpen = IBDir & "ExtForms\" & FullName
	If Not fso.FileExists(FileToOpen) Then FileToOpen = IBDir & Name
	If Not fso.FileExists(FileToOpen) Then FileToOpen = IBDir & "ExtForms\" & Name

	If Not fso.FileExists(FileToOpen) Then FileToOpen = FindFileByIniFile(fso, "" & BinDir & "Config\ExtFormsTree.txt", ".*;(.*);.*", Name)
	'If Not fso.FileExists(FileToOpen) Then FileToOpen = FindFileByIniFile(fso, "" & BinDir & "Config\EFFolders.txt", ".*;(.*);.*", Name)

	If Not fso.FileExists(FileToOpen) Then FileToOpen = FindFile(fso, IBDir, Name, 1) '�� ���?! �� ����� ����������� ����� �� ���� ���������.

	Status ""
	If Not fso.FileExists(FileToOpen) Then Exit Function '������ �� �����. ������.

	Set Doc = Documents.Open(FileToOpen)
	'If Doc Is Nothing Then Exit Function
	TryToOpenFile = true
End Function

' ����������� ����� ����� Name � �������� Path
' ���� ���� ������, �� ������������ ������ ���� � ����, ���� �� ������, �� ������ ������
Function FindFile(fso, Path, Name, Depth)
	Dim Folder, File, SubFolder

	FindFile = ""
	If Depth = RecursiveSearchDepth Then Exit Function

	Status Path
	Set Folder = Nothing
	On Error Resume Next
	Set Folder = fso.GetFolder(Path)
	On Error Goto 0
	If Folder Is Nothing Then Exit Function

	Set Files = Folder.Files
   For Each File in Files
		If UCase(File.Name) = UCase(Name) Then
			FindFile = fso.BuildPath(Path, Name)
			Exit Function
		End If
   Next

	Set Folders = Folder.SubFolders
   For Each SubFolder in Folders
		found = FindFile(fso, fso.BuildPath(Path, SubFolder.Name), Name, Depth + 1)
		If found <> "" Then
			FindFile = found
			Exit Function
		End If
   Next
End Function 'FindFile

' ����� ����� �� ���������, ����������� � �����-�� ini-�����.
' ini-���� ����������� ���������, � ��� �������� ���������� ��� ������ ����������� ���������,
' ��������� � ��������� RE_Pattern. ���������� ��������� ������ ��������� ���� ������, �������� �������
' � ����� ������������� ��� ��� ��������
Function FindFileByIniFile(fso, ini_name, RE_Pattern, Name)
	FindFileByIniFile = ""

	If not fso.FileExists(ini_name) Then Exit Function
	Set ini_file = fso.OpenTextFile(ini_name, 1, False)

	Set RE = New RegExp
	RE.Pattern = RE_Pattern
	RE.IgnoreCase = true
	Do While not ini_file.AtEndOfStream
		line = ini_file.ReadLine
		Set Matches = RE.Execute(line)
		If Matches.Count > 0 Then
			Dir = Matches(0).Submatches(0)
			FindFileByIniFile = FindFile(fso, Dir, Name, 1)
			If FindFileByIniFile <> "" Then Exit Do
		End If
	Loop
	ini_file.Close
End Function 'SearchFileByIniFile

' �-��� �������� ������� ���� ���������� ������
Function TryToOpenClass(WhatToOpen)
	TryToOpenClass = false
	Set ClassInfo = Find1CPPClass(WhatToOpen)
	If ClassInfo Is Nothing Then Exit Function

	Set Doc = ClassInfo.OpenModule()
	If Doc Is Nothing Then Exit Function

	Doc.Open
	Windows.ActiveWnd.Caption = "����� " & ClassInfo.Name
	TryToOpenClass = true
End Function

' �-��� �������� ����� ����� 1�++.
' � ������ ������ ������������ ������ � ����������� � ������
Function Find1CPPClass(ClassName)
	Set Find1CPPClass = Nothing
	If SA Is Nothing Then Exit Function
	
	If SA.ClassCount = 0 Then SA.UpdateClassesInfo
	
	On Error Resume Next
	Set Find1CPPClass = SA.ClassInfo(ClassName)
	On Error Goto 0
End Function

Sub ShowClassesInfo
	SA.UpdateClassesInfo
	Message "ClassCount = " & SA.ClassCount
	for i = 0 to SA.ClassCount-1
		Set class_info = SA.ClassInfo(i)
		Message class_info.Name & ": " & class_info.Location
	next
End Sub

'�������� ������� ����, ��� �������� �������� � ���������
Sub OpenFileFromClipboard()
	sFileName = CommonScripts.GetFileNameFromClipboard
	If sFileName <> "" Then
		Documents.Open sFileName
	End If
	sFileName = ""
End Sub

Function TryGotoTree()
	TryGotoTree = false
	Set Doc = CommonScripts.GetTextDoc(true, 0)
	If Doc Is Nothing Then Exit Function

	CurrDocName = doc.Name
	CurrDocName = split(CurrDocName,".")               
	if UBound(CurrDocName) < 1 Then Exit Function
	
	On Error Resume Next
	Set NavigationTools = scripts("NavigationTools")
	TryGotoTree = NavigationTools.GoToMDTreeItem(CurrDocName(0) & "."  & CurrDocName(1), 0, false)
	On Error Goto 0
	'TryGotoTree = true
End Function

'���� ��� ������� ������ "����\�������" � ��������� �������� ��� �����, �� ������ ��������� ������� ���� ����
Sub Configurator_OnFileDialog(Saved, Caption, Filter, FileName, Answer)
	If Saved Then Exit Sub

	sFileName = CommonScripts.GetFileNameFromClipboard
	If sFileName <> "" Then
		' ���� ������������ ����� ��������� ������ ��� ������ "�������� ��� ����������� ������������"
		bMDWork = false
		bMDFile = false
		If (InStrRev(UCase(sFileName), ".MD") > 0) then
			bMDFile = true
		end if
		If (InStr(UCase(Filter), "*.MD") > 0) then
			bMDWork = true
		end if 

		if bMDWork and  bMDFile then
			If MsgBox("������������ ���� " & sFileName & "?", vbYesNo, SelfScript.Name) = vbYes Then
				FileName  = sFileName
				Answer = mbaOK
			end if

		else
			if not bMDFile and not bMDWork then
				If MsgBox("������� ���� " & sFileName & "?", vbYesNo, SelfScript.Name) = vbYes Then
					' TODO ��� �������� - � ���� ������ �� �������� FileName  = sFileName 
					Documents.Open sFileName
					Answer = mbaCancel
				end if
			end if
		End If
	End If
End Sub

Private Function IsTextWindow()
	IsTextWindow = false

	Set doc = CommonScripts.GetTextDocIfOpened(false, true)

	If Doc Is Nothing Then
		Exit Function
	End If
	IsTextWindow = true
End Function

' ��������� ������ ���� ������� 1�++, ����� �� ������ � �������� ����� ������
' ����� ��������� �������� ������ ������������ ������ ������� ��� ������ �������
Sub SelectAndNavigateToClasses()
	If SA Is Nothing Then
		Message "�� ���������� ��� �� ��������������� ��������� SyntaxAnalysis.wsc", mRedErr
		Exit Sub
	End If

	set Dict = CreateObject("Scripting.Dictionary")

	If SA.ClassCount = 0 Then SA.UpdateClassesInfo
	
	For i = 0 to SA.ClassCount - 1
		Set Info = SA.ClassInfo(i)
		Dict.Add Info.Name, Info.Location
	Next

	Selection = CommonScripts.SelectValue(Dict)
	If Selection = "" Then Exit Sub

	set doc = Documents.Open(Selection)
	If Doc Is Nothing Then Exit Sub

	If Doc.Type = docWorkBook Then
		doc.ActivePage = 1
	end if

	For i = 0 to SA.ClassCount - 1
		Set Info = SA.ClassInfo(i)
		If Info.Location = Selection Then Windows.ActiveWnd.Caption = "����� " & Info.Name
	Next
	
	SendCommand 33298 ' �������� ������ ��������
End Sub ' Sele�tAndNavigateToClasses

Sub UpdateClassesInfo()
	If SA Is Nothing Then
		Message "�� ���������� ��� �� ��������������� ��������� SyntaxAnalysis.wsc", mRedErr
		Exit Sub
	End If
	SA.UpdateClassesInfo()
End Sub
