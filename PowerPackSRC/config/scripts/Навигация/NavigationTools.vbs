$NAME NavigationTools

'==================================================================================================
'	����� �������� ��� ��������� �� ������������
'
' ������: $Revision: 1.22 $
'
'	�����: MetaEditor
'	e-mail: shotfire@inbox.ru
'==================================================================================================

'��������� � ����������:

'	[*]	���������� � ���� ������� GotoFormula, GotoMdTreeItem
'	(����� �������� ��������� ��� ������������� � ���������� � ��.)
'	[*]	����������� ���
'	[*]	����� ��������� ���������� ����� CommonScripts
'	[*]	��������� ��������� ��������� GotoFormula:
'	���� � ������� ������������ "#" �� ���������� ����� ��������������� �������� �����������,
'	���������, �����������, ��� ���������� ����� ������������������� � ���������������������������.
'
'	[*]	#���������������� / ���������� / ��������������� / ��������� �� ������ ���������� -
'	�������� � ����� ���������
'
'	[*]	�������� ������������ �������� ������ � �� ���� � ������� ������ ��������� �� ����������
'	(�� ��������� ��������, ��������������� � ��������� Init)
'	[*] ����� �� ����� �������� � �������� ���������
'
'	[*]	���� ������� ��������� � �������� "��������������.��������������" �� ��� ������� ������ �������
'	� ������������ ���������� ������� � ����.
'	� ����� � ���� ���������� ��������� ��������:
'	���� � ���� ������� �������� ������� ������� �� ������������ �� ������ �������� "���"
'	� � ��� �������� ����� �� ���������� � �����.
'		� �����, ����� ������� ����������� �� ������ ���������������� �������� ���� �������
'	(�������, �������������, ���, ��� �������� (� ������� ������� ��������� ��������� ��������))
'	���� ������ ��� �� ������� ���������� ��� � ������� � ����
'	(��� �����, �� ������ �� �������� "���" ��������� � ������ ��)
'
'	[*]	������ GoToMDTreeItem ��������� ��� ������ �� ����� ������ ����������
'
'	[*]	����� �������� ��������� ������ ��������������� ����� ����� � ���� ������
'	� ������� "�����������" (��������, Ctrl+F ��� �������� ������� ������)

'22.04.05
'	[+] GotoFormula : �������� ����� ������� �� ����������� ������������� ���� ������� � ������� ������.���
'	[*] GotoMDTreeItem :
'	- ���������� �� �������� 1.0.1.9
'	- �������� ������ �������� ShowChilds - ����� �������� ������ ���� �������� ����
'	  ��������� �� ���� ������ (�������� "����������.������������") � ����������
'	  ���������� ����� ��������� (���������, �����, ��������)  - ���������� ������ ���� ����������
'	  ����� ����� �������������� �������
'04.05.05
'	[+] ������� � ���� ����������� �������� � ������ ��
'
'05.05.05
'	[+] ����������� ���� ���������� � ������(SelectInTree) ���������� �������� ��, ����������� � Init - flShowTypes
'
'28.05.05
'	[*] GotoFormula: ��������� ������� c ������� ������ ����������� OpenConf 1.0.2.0
'
'29.05.05
'	[+] ��� �������� ����:
'	������� �����������������(��="",��������="")
'		// ����������� ������� ��������� ���� �������� ��������.
'		// CASE �� ���������� ���������
'		���� ��="" ����� ��=�����.���������������(); ���������;
'
'		���� ��="�������" �����
'			...
'		��������� ��="�����" �����
'			...
'		��������� ��="������������" �����
'			...
'		���������
'	������������
'	��������� ��������� ���������:
'	���� � ������ 15 ������� - �� ������������ ���� ��������� ���� "=�����.���������������()"
'	� ������������� ���������  �������� ����� �������� = "������������"	�����
'	�������������� ������� � ������ ���� "��������� ��="������������""
'	����������� � Init - flTryGotoActiveElementStatement
'
'10.06.05
'	[*] ���������� ���. ��������� ��� ������������� �����, ��������:
'	���� (��="��1") ��� (��="��2") ��� (��="��������������") � (��� = 10) �����
'
'21.06.05
'	[+] ��� ������ ���� "������������("����������.������������",����)"
'	�������������� ������� � ������ �� � "����������.������������"
'	[+]� ������� SelectMetadataAndGotoObj ����� �� ������ ������� "�����.������" ��� ��� �� �������,
'      ��������� ��������, ������� �� ���� ������������, ���� ���, �� �����������.
'
'18.07.05
'	[+] ����� ������ GotoControlWithFormula, �������� GotoFormula
'	���� ����� ��� �������� � ���� ��� �������� ��� � ������� ���� ��� �����
'	����� ������ ���� ���������, � ��� ������ ������������ ��� �� �����
'==================================================================================================


Const WM_CLOSE = &H10
Const WM_GETTEXT = &HD
Const WM_COMMAND =  &H111
Const WM_LBUTTONDOWN =  &H0201
Const WM_SETFOCUS = &H0007

Const BM_CLICK = &H00F5

Const TVM_GETNEXTITEM = 4362
Const TVM_SELECTITEM  = 4363
Const TVM_EXPAND      = 4354

Const TVE_EXPAND = 2
Const TVE_TOGGLE = 3

Const TVGN_ROOT =  0
Const TVGN_NEXT =  1
Const TVGN_CHILD = 4
Const TVGN_CARET = 9

Const cmdOpenCfgWindow   = 33188 '������� ���� ������������
Const cmdOpenPropWindow  = 32844 '������� ���� ������� ��������
Const cmdOpenEditWindow  = 33711 '������� ���� �������������� ��������
Const cmdEditDescription = 33718 '������� ���� �������������� �������� ��������
Const cmdSaveAsExternal  = 33239 '������� ������ "��������� ��� �������..."(��� ������� � ���������)

Dim flSearchInGM '���� ��������������� ������ � �� (����������� � ��������� Init)
Dim flShowTypes  '���������� ���� ���������� (����������� � ��������� Init)
Dim flTryGotoActiveElementStatement '������ � ���� ��������� "�����.���������������" (����������� � ��������� Init)

Dim Wrapper
Dim MDDict

'=========================================================================================
'GOTOFORMULA
'==========================================================================================
Function GoToFormula()
	GoToFormula = false
	Formula = GetFormula()
	Identifier = GetActiveCtrlIdentifier()
	'message "ActiveCtrlIdentifier: " & Identifier
	if Formula="" then Exit Function

	'stop
	if GotoMdTreeItem (Formula, 0, false) then
  		GoToFormula = true
  		exit Function
  	end if

  	if TryToOpenForm (Formula) then
  		GoToFormula = true
  		exit Function
  	end if

	set doc = CommonScripts.GetTextDocIfOpened(1)
	if doc is nothing then exit Function

	If Instr(Formula,"#") > 0 then
 		list = "�����������()" & vbCrLf & "�����������()" 	& vbCrLf & "���������()"
 		if Instr(doc.Name,"��������") = 1 then
 			list = list & vbCrLf & "�������������������()" & vbCrLf & "���������������������������()"
 		end if
		Formula = CommonScripts.SelectValue(list)
 		if Formula = "" then exit Function

 		if Formula = "�������������������()" or Formula = "���������������������������()" then
			CurrDocName = split(doc.Name,".")
			if not IsArray(CurrDocName) then exit Function
			if UBound(CurrDocName) < 1 then exit Function
			set doc = Documents(CurrDocName(0) & "."  & CurrDocName(1) & ".������ ���������")
			doc.Open
 		end if
	End If

	if Instr(Formula,"(") = 0 then Formula = Formula & "()" 'Exit Sub

	FirstLine = doc.Range(0)
	If UCase(left(FirstLine,18)) = "#���������������� " then
		Name = trim(mid(FirstLine,19))
    	If left(name,2)="\\" or mid(name,2,1)=":" Then
			On Error Resume Next
			set doc = Documents.Open(Name) '���������� ����
			On Error GoTo 0
			if doc is nothing then message "������ �������� ����� �� ��������� #����������������" : exit Function
		Else
			On Error Resume Next
			set doc = Documents.Open (IBDir & Name) ' ���� ����������� �������� ��
			On Error GoTo 0
			if doc is nothing then message "������ �������� ����� �� ��������� #����������������" : exit Function
		End If
	End If

	If Left(UCase(Formula),2) = "��" Then ' ���� ��� ���������� ���������/�������
		set doc = Documents("����������������")
		doc.Open
	End If

    Formula = Left(Formula,Instr(Formula,"(")-1) '��� ���������/������� ��� ������

	If Instr(Formula,"[") > 0 Then '����� ���� "[" (���� ���, ��������, ������ � ������)
		Formula = Trim(Replace(Formula,"[",""))
	End If

	if Formula="" then Exit Function

	Status "����� ��������� " & Formula & " ..."
	pos = CommonScripts.FindProc(doc,Formula)
	Status ""
	if pos > -1 then
		doc.MoveCaret pos, 0, pos, doc.LineLen(pos)
		TryGotoActiveElementStatement doc,Identifier
	else
		if flSearchInGM then
			set doc1 = Documents("����������������")
			Status "�������������� ����� ��������� " & Formula & "� ���������� ������..."
			pos1 = CommonScripts.FindProc(doc1,Formula)
			Status ""
			if pos1 > -1 then
				doc1.Open
				doc1.MoveCaret pos1, 0, pos1, doc1.LineLen(pos1)
				GoToFormula = true
				exit Function
			end if
		end if
		CreateProcedure doc,Formula
	end if
	GoToFormula = true
End Function

'==========================================================================================
Private Function GetActiveCtrlIdentifier()
	GetActiveCtrlIdentifier = ""
	set Form = Windows.ActiveWnd.Document.Page(0)
	Identifier = Form.ctrlProp(cInt(Form.Selection),cpStrID)
	GetActiveCtrlIdentifier = Identifier
End Function

'==========================================================================================
Private Function TryToOpenForm(Formula)
	TryToOpenForm = false
	'Formula = "������������(""�����.���������������_"",��������)"
	'patt = "((������������\("")|(openform\(""))([\w�-���\d]+\.[\w�-���\d]+)"""
	patt = "������������\(""([\w�-���\d]+\.[\w�-���\d]+)"""
	set Matches = CommonScripts.RegExpExecute(patt,Formula)
	if not Matches is Nothing then
		For Each Match In Matches
			'message Match.SubMatches(0)
			GotoMdTreeItem Match.SubMatches(0), 0, false
			TryToOpenForm = true
			exit for
		Next
	end if
End Function

'==========================================================================================
Private Function TryGotoActiveElementStatement(doc,Identifier)
	if not flTryGotoActiveElementStatement then exit function
	if Identifier = "" then exit function
	cnt = 0
	l = doc.SelStartLine + 1

	patt = "([\w�-���\d]+)\s*=\s*�����\.���������������\(\)"
	while cnt < 15
		sLine = doc.Range(l)
		if CommonScripts.RegExpTest("��������������|������������",sLine) then exit function
		if CommonScripts.RegExpTest("^(\s*)[\r\n]*$|^(\s*//.*)",sLine) then
			'message "sLine " & sLine
			l=l+1
		else
			set Matches = CommonScripts.RegExpExecute(patt, sLine)
			'message "try to find active element statement"
			'message "sLine " & sLine
			if not Matches is Nothing then
				For Each Match In Matches
					param = Match.SubMatches(0)
					exit for
				Next
				'message "found param " & param
				'message "in string " & sLine
				cnt = 16
			else
				'message "param not found"
				cnt=cnt+1
			end if
		end if
	Wend

	for i = l to doc.LineCount
		sLine = doc.Range(i)
	    if CommonScripts.RegExpTest("��������������|������������",sLine) then exit function
		patt = "����\s+[ -�]*\(*" & param & "\s*=\s*""" & Identifier & """[ -�]*\s+�����"
		if CommonScripts.RegExpTest(patt,sLine) Then
			'message sLine
			doc.MoveCaret i, 0, i, doc.LineLen(i)
			exit for
		end if
	next
End Function

'==========================================================================================
Private Function GetFormula()
	GetFormula = ""
	Title = Space(20)
	Formula = Space(128)
	ForegroundWnd = Wrapper.GetForegroundWindow

 	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
	cnt = Wrapper.SendMessage(ForegroundWnd, WM_GETTEXT ,20,Title) ' (��������� ����)

	Title = UCase(CStr(Title))
	If Left(Title,8)="��������" Then '���� ��� ���� �������
    	ActiveControl = Wrapper.GetFocus '������� ���� ������� � ������� �����

		Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
		TabControl = Wrapper.FindWindowExA (ForegroundWnd,0,"SysTabControl32",NULL) 'TabControl � ����������

		'���� ������� ����� ����� �� ��������� ����� �� ���� ����� �������� ����� �������� "���"
		if ActiveControl = ForegroundWnd or ActiveControl = TabControl then
    		Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llls", "f=s", "r=l"
			TypeTab = Wrapper.FindWindowExA(ForegroundWnd,0,NULL,"��� ") '������ �������� "���"

			if TypeTab = 0 then
				Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	  			Wrapper.SendMessage ForegroundWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
  				Exit Function
			end if
			Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
			TypeCBBox = Wrapper.FindWindowExA(TypeTab,0,"COMBOBOX",NULL)
			if TypeCBBox = 0 then
				Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
	  			Wrapper.SendMessage ForegroundWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
  				Exit Function
			end if
			cnt = Wrapper.SendMessage (TypeCBBox, WM_GETTEXT ,128, Formula) '������� ����� �� ����������
		else
			cnt = Wrapper.SendMessage (ActiveControl, WM_GETTEXT ,128, Formula) '  ��������� �������
		end if

		GetFormula =  Trim(Cstr(Formula))

		Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llss", "f=s", "r=l"
	  	OKButtonHandle = Wrapper.FindWindowExA(ForegroundWnd,0,"Button","��������") '����� ������ "��������"
	  	Wrapper.SendMessage OKButtonHandle, BM_CLICK ,0 ,0 '����� �� ��
	 	Wrapper.SendMessage ForegroundWnd, WM_CLOSE ,NULL, NULL '������� ���� �������

	ElseIf ForegroundWnd=Windows.MainWnd.HWND Then '���� ��� ���� �������������
		If Windows.ActiveWnd.Document <> docWorkBook Then Exit Function
		If Windows.ActiveWnd.Document.ActivePage <> 0 Then Exit Function

		If Version >= 1020 Then
			set Form = Windows.ActiveWnd.Document.Page(0)
			Formula = Form.ctrlProp(cInt(Form.Selection),cpFormul)
			'���� � ������� ������ ���, ������� ����� ������������� ��������
			if Trim(Formula) = "" then Formula = GetActiveCtrlIdentifier()
		end if
		
		if Trim(CStr(Formula)) = "" then
			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
			Wrapper.SendMessage Windows.MainWnd.HWND, WM_COMMAND ,cmdOpenPropWindow, NULL '������� �������� ��������

			PropWnd = Wrapper.GetForegroundWindow '���� ������� ��������

			if PropWnd = Windows.MainWnd.HWND then '���� ������� ������� �� ���������
				Set WinFinder = CreateObject("ArtWin.Win")
				PropWnd = WinFinder.Find("�������� ")
				Wrapper.Register "USER32.DLL",   "GetClassName",    "I=lrl", "f=s", "r=l"
				ClassName = Space(128)
				cnt = Wrapper.GetClassName(PropWnd, ClassName, 128)
				if Left(cstr(ClassName), cnt) <> "#32770" then PropWnd = 0
			end if

			if PropWnd = 0 then Exit Function

			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
			cnt = Wrapper.SendMessage(PropWnd, WM_GETTEXT ,20,Title) ' (��������� ����)
			Title = UCase(CStr(Title))

			If Left(UCase(Title),8) <> "��������" then
				Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
		  		Wrapper.SendMessage PropWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
		  		Exit Function
	  		End If

			Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
			TabControl = Wrapper.FindWindowExA (PropWnd,0,"SysTabControl32",NULL) '����� TabControl � ����������

			'����� �� "�������������" (������ ���������� �������� � ����. ������)
			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
			Wrapper.SendMessage TabControl,WM_LBUTTONDOWN,0,cLng(125 + 10 * &H10000)

			Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llls", "f=s", "r=l"
			DopTab = Wrapper.FindWindowExA(PropWnd,0,NULL,"�������������") '������ �������� "�������������"

			if DopTab = 0 then '��� �������� "�������������"
				Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
		  		Wrapper.SendMessage PropWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
				Exit Function
			end if

			Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
			FormulaEdit = Wrapper.FindWindowExA(DopTab,0,"EDIT",NULL)	'����� ���� ����� �������

			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
			cnt = Wrapper.SendMessage (FormulaEdit, WM_GETTEXT ,128, Formula) '������� ����� �������

	    	if Trim(Cstr(Formula)) = "" then
				Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
		  		Wrapper.SendMessage PropWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
				Exit Function
			End If

			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
		  	Wrapper.SendMessage PropWnd, WM_CLOSE ,NULL, NULL '������� ���� �������
	  	'Else
		' 	set Form = Windows.ActiveWnd.Document.Page(0)
		'	Formula = Form.ctrlProp(cInt(Form.Selection),cpFormul)
	  	End If

  		GetFormula = Trim(CStr(Formula))
	End If
End Function

'========================================================================
Private Sub CreateProcedure(Doc,Formula)
	ProcOrFunc = InputBox("��������� ��� �������" & vbCrLF & Formula & " - �� ����������." & vbCrLF & vbCrLF & _
	"��� ���������: 1 - ���������, 2 - ������� ", "������� ��������", "1")

	If ProcOrFunc="1" then
		ReplValue1 = "��������� " : ReplValue2 = "�������������� // "
	ElseIf ProcOrFunc="2" then
		ReplValue1 = "������� " : ReplValue2 = "������������ // "
	Else
		Exit Sub
	End If

	ModuleText = split(doc.Text, vbCrLf)
	If not IsArray(ModuleText) then exit sub
	if UBound(ModuleText) > 0 Then
		doc.MoveCaret UBound(ModuleText), 0 '��������� � ����� ������
		for i = doc.SelStartLine to 0 step -1 '����� ��������� ���������/�������
			sText = lTrim(UCase(ModuleText(i)))
			if Instr(sText,"��������������") = 1 or Instr(sText,"������������") = 1 then
				doc.MoveCaret i, 0
				Exit For
			end if
		next

		If i+1=doc.LineCount then '���� ����� ������, �� ������� ������
			doc.Range(i,Len(sText),i,Len(sText))=vbCrLF
			doc.MoveCaret i+1, 0
		Else
			doc.MoveCaret i+1, 0
		End If
	End If

	Text=vbCrLF & "//" & String(70,"=") & vbCrLF & _
	ReplValue1 & Formula & "()" & vbCrLF & vbTab & vbCrLF & _
	ReplValue2 & Formula & "()" & vbCrLF
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = Text
	doc.MoveCaret i+4, 1
End Sub

'========================================================================
'GOTOMDTREEITEM
'========================================================================
'����������.������������
Private Function BuildTree(MDO,level)
	txt = ""
	For i = 1 To level
    	txt = txt & vbTab
	Next
	if flShowTypes then
		on error resume next
		sType = " [" & mdo.Type.Fullname & "]"
		on error goto 0
		if sType = " []" then sType = ""
	end if
	txt =  txt + mdo.Name & sType & "|e" & vbCrLf
	Set Childs = mdo.Childs
	For i = 0 To Childs.Count - 1
		txt = txt & vbTab & Childs.Name(i, True) & vbCrLf
		Set mda = Childs(i)
		For j = 0 To mda.Count - 1
			txt = txt + BuildTree (mda(j), level + 2)
		Next
	Next
	BuildTree = txt
End Function

'========================================================================
Private Function PrepareMDPath(Path) ', flDeleteInvisibleChilds
	PrepareMDPath = Path

	PathParts = Split(Path,"\")

	ItemType = UCase(PathParts(0))

	If ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "����������"     then
		PathParts(0) = "�����������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(2) = Replace(PathParts(2),"��������","#")
			'end if
		end if

	ElseIf	ItemType = "��������"       then
		PathParts(0) = "���������"
		if UBound(PathParts) > 1 then
			PathParts(2) = Replace(PathParts(2),"�������������","�����")
			PathParts(2) = Replace(PathParts(2),"����������������������","��������� �����")
		end if
	ElseIf	ItemType = "������"         then
		PathParts(0) = "������� ����������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(2) = Replace(PathParts(2),"�����","#")
			'end if
		end if

	ElseIf ItemType = "������������"   then
		PathParts(0) = "������������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(2) = Replace(PathParts(2),"��������","#")
			'end if
		end if

	ElseIf ItemType = "�����"          then
		PathParts(0) = "������"

	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "����������"     then
		PathParts(0) = "����� ������"

	ElseIf ItemType = "�����������"    then
		PathParts(0) = "���� ��������"

	ElseIf ItemType = "��������"       then
		PathParts(0) = "��������"

	ElseIf ItemType = "��������"       then
		PathParts(0) = "��������"

	ElseIf ItemType = "�������"        then
		PathParts(0) = "��������"
		if UBound(PathParts) > 1 then
			PathParts(2) = Replace(PathParts(2),"���������","���������")
			PathParts(2) = Replace(PathParts(2),"������","�������")
			PathParts(2) = Replace(PathParts(2),"��������","���������")
		end if

	ElseIf ItemType = "��������������" then
		PathParts(0) = "������� ��������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(2) = Replace(PathParts(2),"��������","#")
			'end if
		end if

	ElseIf ItemType = "����������"     then
		PathParts(0) = "���� ��������"

	ElseIf ItemType = "��������������"  then
		PathParts(0) = "������ ��������"

	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������\����������"

	ElseIf ItemType = "������������������"      then
		PathParts(0) = "���������\������������������"

	ElseIf ItemType = "����������������������"      then
		PathParts(0) = "���������\����� ���������"

	'   --------------
	'- �������� ������� -
	'   --------------
	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "�����������"     then
		PathParts(0) = "����������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(1) = PathParts(1) & "\��������"
			'end if
		end if

	ElseIf ItemType = "���������"       then
		PathParts(0) = "��������"
		if UBound(PathParts) > 1 then
			PathParts(2) = Replace(PathParts(2),"�����","�������������")
			PathParts(2) = Replace(PathParts(2),"��������� �����","����������������������")
		end if

	ElseIf ItemType = "������� ����������" then
		PathParts(0) = "������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(1) = PathParts(1) & "\�����"
			'end if
		end if

	ElseIf ItemType = "������������"   then
		PathParts(0) = "������������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(1) = PathParts(1) & "\��������"
			'end if
		end if

	ElseIf ItemType = "������"          then
		PathParts(0) = "�����"

	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "����� ������"     then
		PathParts(0) = "����������"

	ElseIf ItemType = "���� ��������"    then
		PathParts(0) = "�����������"

	ElseIf ItemType = "��������"       then
		PathParts(0) = "��������"

	ElseIf ItemType = "��������"       then
		PathParts(0) = "��������"

	ElseIf ItemType = "��������"        then
		PathParts(0) = "�������"
		if UBound(PathParts) > 1 then
			PathParts(2) = Replace(PathParts(2),"���������","���������")
			PathParts(2) = Replace(PathParts(2),"�������","������")
			PathParts(2) = Replace(PathParts(2),"���������","��������")
		end if

	ElseIf ItemType = "������� ��������" then
		PathParts(0) = "��������������"
		if UBound(PathParts) > 1 then
			'if flDeleteInvisibleChilds = true then
				PathParts(1) = PathParts(1) & "\��������"
			'end if
		end if

	ElseIf ItemType = "���� ��������"     then
		PathParts(0) = "����������"

	ElseIf ItemType = "������ ��������"  then
		PathParts(0) = "��������������"

	ElseIf ItemType = "���������"      then
		PathParts(0) = "���������"

	ElseIf ItemType = "����������"      then
		PathParts(0) = "���������"
	End If

	If UBound(PathParts) > 0 then
		If PathParts(1) = "������������������"  then
			PathParts(1) = "������������������"
		ElseIf PathParts(1) = "����� ���������"      then
			PathParts(1) = "����������������������"
		End If
	End If

	If UBound(PathParts) = 0 then
		Path = PathParts(0)
	else
		Path = Join(PathParts,"\")
	end if

	Path =  Replace(Path,"\#\","\")
	PrepareMDPath = Path
End Function

'========================================================================
'�����
'���������.�����������������
'����������.������������
'����������.���������������
'��������.���
'������.�����������
'������������.������������
'�����.������������
'���������.����������������
'����.��������
'����������.��������
'�����������.��������������������
'�������.�����
'�������.������
'��������������.��������
'����������.������������������05
'��������������.������������������������������������
'���������.�������12
'========================================================================
Function GoToMDTreeItem(ByVal Path,Action,ShowChilds)
	GoToMDTreeItem = false
	if Version < 1019 then
		message "��������� ������ ��������� �� ����� 1.0.1.9"
		exit Function
	end if

	Path = Replace(Path,".","\")

	if Right(Path,1) = "\" then Path = Left(Path,Len(Path)-1)

	ItemType = ""
	ItemKind = ""

	if InStr(Path,"\") > 0 then
		PathParts = Split(Path,"\")
		ItemType = UCase(PathParts(0))
		ItemKind = UCase(PathParts(1))
	else
		ItemType = Path
	end if

	if ItemType = "����" then ItemType = "����������"


	set tmp = nothing
	On Error Resume Next
	if ItemKind = "" then
		Set tmp = MetaData.TaskDef.Childs(CStr(ItemType))
	else
		Set tmp = MetaData.TaskDef.Childs(CStr(ItemType))(CStr(ItemKind))
	end if
	On Error Goto 0
	if tmp is nothing then	exit Function


	if ShowChilds = true then
		if IsArray(PathParts) then
			if UBound(PathParts) = 1 then
				if Instr("����������,��������,������,������������,�������,��������������",ItemType) <> 0 then
					'� ���� �������� ���� ���������
					Set Srv=CreateObject("Svcsvc.Service")
					Vals =  BuildTree(MetaData.TaskDef.Childs(CStr(ItemType))(CStr(ItemKind)),0)
					Vals = ItemType & "\" & Vals
					if Instr("����������,������,��������������",ItemType) <> 0 then
						pos = Instr(Vals,"�����������") '����� �� ������ ����� �������
						Vals = Left(Vals,pos-1)
					end if
					Path = srv.SelectInTree(Vals,"",False,False)
					if Instr(Path,"[") > 0 then	Path = Left(Path,Instr(Path,"[")-1)
					if Path = "" then
						GoToMDTreeItem = true
						Exit Function
					end if
				end if
			end if
		end if
	end if

	Path = PrepareMDPath(Path) ',true
	MDWnd.DoAction Path,Action

	GoToMDTreeItem = true
End Function


'========================================================================================
Function GoToSelectedMDTreeItemType()
	GoToSelectedMDTreeItemType = false
	if Instr(Windows.ActiveWnd.Caption,"������������")=1 then
		if MDWnd.ActiveTab = 0 then
			path = MDWnd.GetSelected()
			path = PrepareMDPath(path)
			parts = Split(Path,"\")
			on error resume next
			set mdo = MetaData.TaskDef.Childs(CStr(Parts(0)))(CStr(Parts(1))).Childs(CStr(Parts(2)))(CStr(Parts(3)))
			itemtype = mdo.Props("���") & "." & mdo.Props("���")
			if itemtype = vbEmpty then
				Err.Clear
				set mdo = MetaData.TaskDef.Childs(CStr(Parts(0)))(CStr(Parts(1)))
				itemtype = mdo.Props("���") & "." & mdo.Props("���")
			end if
			if Err.Number = 0 then
				GotoMDTreeItem itemtype, 0, false
				GoToSelectedMDTreeItemType = true
			end if
			on error goto 0
		end if
	end if
End Function

'========================================================================================
' ����� ������� ����������� �� ������ ���������� � ������� � ����
'========================================================================================
Sub SelectMetadataAndGotoObj()
	on Error resume next
	Set childs = MetaData.TaskDef.childs
	
	For i = childs.Count - 1 To 0 Step -1
		Set mda = childs(i)
		For j = mda.Count - 1 To 0 Step -1
			lname = mda(j).Name
			If Err.Number = 0	then
				'vals = vals & vbCrLf & childs.Name(i, True) & "." &  mda(j).Name
				str = childs.Name(i, True) & "." &  mda(j).Name
				If not MDDict.Exists(str) then MDDict.Add str, ""
			else
				Err.Clear
			end if
		Next
	Next

	Scripts("ChangeKbdLayout").ActivateRusLayout()
	on Error goto 0
	
	vals = ""                             
	keys = MDDict.Keys
	for i = MDDict.Count - 1 to 0 step -1
		vals = vals & keys(i) & vbCrLf
	next
	
	fStr = CommonScripts.SelectValue(Vals)
	if fStr = "" then Exit Sub
	
	MDDict.Remove fStr
	MDDict.Add fStr, ""
		
	'����� �� "���������"
	WhatToOpen = "�����.������"
	NameToOpen = fStr
	ListForms = WhatToOpen
	Str = NameToOpen
	ObjType = Mid(fStr,1,Instr(fstr,".")-1)
	ObjName = Mid(fStr,Instr(fstr,".")+1)

	If ObjType = "����������" or ObjType = "������" Then
		ListForms = Scripts("���������").GetListForms(ObjType, ObjName, WhatToOpen)
		If ObjType = "����������" Then
			ListForms = WhatToOpen & "," & Replace(WhatToOpen, "�����.", "����� ������.") & "," & ListForms
		ElseIf ObjType = "������" Then
			ListForms = ListForms
		Else
			ListForms = WhatToOpen & "," & ListForms
		End If
	ElseIf ObjType = "��������" Then
		ListForms = ListForms & "," & "������ ���������"
	End If

	'If Instr("������������,�������",ObjType) > 0 Then
	'	ListForms = Replace(ListForms,"�����.������","")
	'End If

	Set Doc = Nothing
	On Error Resume Next
	Set Doc = Documents(ObjType & "." & ObjName & ".�����.������")
	On Error Goto 0
	if doc is Nothing then
		ListForms = Replace(ListForms,"�����.������","")
	end if

	List = ListForms & ",����� � ������,��������,�������������"
	List = Replace(List, ",", vbCrLf)

	WhatToOpen = CommonScripts.SelectValue(List, fStr, vbCrLf, true)
	If WhatToOpen <> "" Then
		If not IsConfigWndOpen Then
			Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
			Wrapper.SendMessage Windows.MainWnd.HWND, WM_COMMAND ,cmdOpenCfgWindow, NULL
		end if
		if WhatToOpen = "����� � ������" then
			GoToMDTreeItem fStr, 0, false

		elseif	WhatToOpen = "��������" then
			GoToMDTreeItem fStr, mdaProps, false

		elseif	WhatToOpen = "�������������" then
			GoToMDTreeItem fStr, cmdOpenEditWindow, false

		else ModuleName = NameToOpen & "." & WhatToOpen
			Documents(ModuleName).Open
			Set Doc = Documents(ModuleName)
			If Not Doc Is Nothing Then
				If Doc.IsOpen Then
					nLine = CommonScripts.FindProc(Doc, "�����������", "���������","�������������������")
					If nLine < 0 Then nLine = 0
					CommonScripts.Jump nLine, -1, -1, -1, ModuleName
				End If
			End If
		end if
	End If
End Sub ' SelectMetadataAndGotoObj

'========================================================================================
'ACTIVATEFINDEDIT  (Ctrl + F)
'========================================================================================
Sub ActivateFindEdit()
	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"

	combo = Wrapper.FindWindowExA(Windows.MainWnd.Hwnd,0,"AfxControlBar42",NULL)
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llss", "f=s", "r=l"
	combo = Wrapper.FindWindowExA(combo,0,"Afx:400000:8:10011:10:0","�����������")
	combo = Wrapper.FindWindowExA(combo,0,"ComboBox","")
	Wrapper.SendMessage combo,WM_SETFOCUS,0,0
End Sub

'========================================================================================
'GOTOCONTROLWITHFORMULA
'========================================================================================
Sub GotoControlWithFormula()
	If Windows.ActiveWnd.Document <> docWorkBook Then Exit Sub
	If Windows.ActiveWnd.Document.ActivePage <> 1 Then Exit Sub 
	
	set Form = Windows.ActiveWnd.Document.Page(0)
	set doc = Windows.ActiveWnd.Document.Page(1)
	CurrWord = doc.CurrentWord
	
	if CurrWord = "" then exit Sub
	Set FrmDict = CreateObject("Scripting.Dictionary")
	for i = 0 to Form.ctrlCount - 1
		Formula = Form.ctrlProp(i,cpFormul)

		if Instr(Formula,CurrWord) > 0 then
			Type_ = Form.ctrlType(i)
			Capt = Form.ctrlProp(i,cpTitle)
			strID = Form.ctrlProp(i,cpStrID)
			ControlString = Trim(cStr(i) & " - " & Type_ & ": " & Capt & " ��: " & strID)
			FrmDict.Add ControlString, i
		End If
	next
	
	if FrmDict.Count = 0 then
		msgbox "������� """ & CurrWord & """ �� ������������� � ��������� ���������� �����!"
		Exit Sub
	end if
	
	if FrmDict.Count > 1 then
		ControlPos = CommonScripts.SelectValue(FrmDict)
	else
		ControlPos = FrmDict.Item(ControlString)
	end if
	
	if ControlPos = "" then exit Sub
	
	Windows.ActiveWnd.Document.ActivePage = 0
	
	Form.LayerVisible(Form.ctrlProp(ControlPos,cpLayer)) = true

	Form.Selection = cStr(ControlPos)
End Sub

'========================================================================================
'GotoControlOrFormula
'
' � ����������� �� ��������� (����� ��� ������) ����������� �������
'		- ���� ������� �����(������),  ����������� �������
'				�� ������� �������� �������� (GotoFormula)
'		- ���� ������� ������,  ����������� ������� �� �������,
'				� ������� �������� ���� ������� ��������� (GotoControlWithFormula)
'========================================================================================
Sub GotoControlOrFormula()
	If Windows.ActiveWnd.Document <> docWorkBook Then Exit Sub

	If Windows.ActiveWnd.Document.ActivePage = 0 Then
		GotoFormula

	ElseIf Windows.ActiveWnd.Document.ActivePage = 1 Then
		GotoControlWithFormula

	end if
End Sub

'========================================================================================
'INIT
'========================================================================================
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

	Set Wrapper = CreateObject("DynamicWrapper")

	Wrapper.Register "USER32.DLL",   "GetFocus",                 "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetForegroundWindow",      "f=s", "r=l"
	Wrapper.Register "USER32.DLL",   "GetParent",         "I=l", "f=s", "r=l"

	flSearchInGM = true 'false/true
	flShowTypes = true 'false/true
	flTryGotoActiveElementStatement = true 'false/true
	Set MDDict = CreateObject("Scripting.Dictionary")
End Sub

'========================================================================================
Init '��� �������� ������� ��������� �������������