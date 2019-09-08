'-------------------------------------------------------------------------------
Function GetDefaultStringArr(CtrlType)
    Select Case CtrlType
    Case "Button"
        GetDefaultString = "{""������"",""BUTTON"",""1342177291"",""8"",""22"",""41"",""13"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "CheckBox"
        GetDefaultString = "{""�����"",""CHECKBOX"",""1342177283"",""7"",""39"",""43"",""13"",""0"",""0"",""4159"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "Switch"
        GetDefaultString = "{""�����"",""RADIO"",""1342177289"",""7"",""54"",""48"",""13"",""0"",""0"",""4158"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "List"
        GetDefaultString = "{"""",""LISTBOX"",""1352663297"",""22"",""40"",""220"",""90"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "Combo"
        GetDefaultString = "{"""",""COMBOBOX"",""1352663107"",""7"",""87"",""55"",""15"",""0"",""0"",""4266"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "Frame"
        GetDefaultString = "{""����� ������"",""1CGROUPBOX"",""1342177287"",""7"",""105"",""58"",""16"",""0"",""0"",""4155"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "Label"
        GetDefaultString = "{""�����"",""STATIC"",""1342177280"",""7"",""126"",""63"",""13"",""0"",""0"",""4154"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""192"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""400"",""0"",""0"",""0"",""204"",""1"",""2"",""1"",""34"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "Text"
        GetDefaultString = "{"""",""1CEDIT"",""1350565888"",""7"",""144"",""66"",""13"",""0"",""0"",""4153"","""","""","""",""-1"",""S"",""10"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "VT"
        GetDefaultString = "{"""",""TABLE"",""1352663040"",""22"",""40"",""220"",""90"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""��������"",""{""""0"""",""""0""""}""}"
    Case "TableColumn"
        GetDefaultString = "{""4"",""�����"",""70"",""STATIC"",""4158"","""","""","""",""0"",""2"",""10"",""0"",""0"",""0"",""2"","""",""0"",""268451840"","""","""","""",""0""},"
    End Select   
	
	GetDefaultStringArr = Split(Mid(GetDefaultString, 3), """,""")
End Function 
	
'_________________________________________________________________________________
Function GetSinonimFromID(Ident) 
	
	Sinonim = Left(Ident, 1)
	for i = 2 to len(Ident) - 1
		NextLett = Mid(Ident, i, 1)
		if LCase(NextLett) = NextLett then
			Sinonim = Sinonim + NextLett
		else
			PreLetter  = Mid(Ident, i - 1, 1) 
			PostLetter  = Mid(Ident, i + 1, 1)  
			PostPostLetter  = Mid(Ident + "#", i + 2, 1) 
			if LCase(PreLetter) = PreLetter Or LCase(PostLetter) = PostLetter then
				Sinonim = Sinonim + " "
			end if
			
			if LCase(PostLetter) = PostLetter then
				NextLett = LCase(NextLett)
			elseIf LCase(PreLetter) = PreLetter and LCase(PostPostLetter) = PostPostLetter  and PostPostLetter <> "#" then
				NextLett = LCase(NextLett)
			end if  
			Sinonim = Sinonim + NextLett
		end if
	Next
	
	Sinonim = Sinonim + Right(Ident, 1)  
	GetSinonimFromID = Sinonim
	
End Function
		
'---------------------------------------------------------------------------------
function AddProcsToModuleBegin(ModuleArr, ProcBlock)  
	
	FlagFound = False
	StringNum = 0
	for i = 0 to UBound(ModuleArr)
		NextLine = UCase(ModuleArr(i))
		if (instr(NextLine, "���������") > 0 Or instr(NextLine, "�������") > 0 ) And _
			instr(NextLine, " �����") = 0 then
				for j = i-1 to 0 step - 1
					PrevLine = Trim(ModuleArr(j))
					if Left(PrevLine, 2) <> "//" And PrevLine <> "" then
						FlagFound = True
						StringNum = j + 1
						exit for
					end if
				Next
				exit for
		end if
	Next
	
	if Not FlagFound then StringNum = 0
	ModuleArr(StringNum) = ProcBlock + ModuleArr(StringNum)   
	
	AddProcsToModuleBegin = 1
end function   
	
'---------------------------------------------------------------------------------
function SetType(ControlArr, ObjType) 
	
	If Instr(ObjType, ".") = 0 then
		ObjKind = ""
		ObjTypeMain = ObjType
	else
		ObjKind = Mid(ObjType, Instr(ObjType, ".") + 1)
		ObjTypeMain = Left(ObjType, Instr(ObjType, ".") - 1)  
	end if 
	
	ObjTypeMain = UCase(ObjTypeMain)
	
	cntType = "U"
	cntKind = "0"
	cntLength = "0"
	cntTochn = "0" 
	NumberMD = -1
	CntView = "1CEDIT"
	
	If Instr(ObjTypeMain, "�����") = 1 then
		Params = Split(ObjTypeMain, "-")
		cntType = "N"
		if UBound(Params) > 0 then
			cntLength = Params(1)
		end if
		if UBound(Params) > 1 then
			cntTochn = Params(2)
		end if 
		
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "������") = 1 then
		Params = Split(ObjTypeMain, "-")
		cntType = "S"
		if UBound(Params) > 0 then
			cntLength = Params(1)
		end if
		
	elseIf Instr(ObjTypeMain, "����") = 1 then

		cntType = "D"
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "����") = 1 then

		cntType = "T"
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "����������") = 1 then
	
	    cntType = "B" 
		NumberMD = 1  
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "��������") = 1 then
	
	    cntType = "O" 
		NumberMD = 2
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "������������") = 1 then
	
	    cntType = "E" 
		NumberMD = 4 
		CntView = "BMASKED"
		
	end if  
	
	If NumberMD > 0 And ObjKind <> "" then
		ObjKind = UCase(ObjKind)  
		Set MainObj = MetaData.TaskDef  
		Set ObjArray = MainObj.Childs(NumberMD)   
		for i = 0 to ObjArray.Count - 1
			if UCase(ObjArray(i).Name) = ObjKind then
				cntKind = ObjArray(i).ID
				Exit for
			end if
		Next
	end if   
	
	ControlArr(1) = CntView
	ControlArr(14) = CntType
	ControlArr(15) = CntLength
	ControlArr(16) = CntTochn
	ControlArr(17) = CntKind
	
	SetType = 1
	
end function       
	
'---------------------------------------------------------------------------------
Sub PoleVvodaSKnopkami()
	
	Set doc = Windows.ActiveWnd.Document.Page(0)
	DlgText = doc.Stream  
	
	Set ModuleDoc = Windows.ActiveWnd.Document.Page(1)  
	ModuleText = ModuleDoc.Text
	
	Ident = inputBox("������� ������������� ���� �����")
	ObjType = inputBox("������� ��� ��������. ����������� ����� ������� � ���� ���.�����������, " + _
		"��������� - ���.����, ������������ - ���.����������������, ����� - �����-15-2, ������ - ������-100")
		
	if left(UCase(ObjType), 4) = "���." then
		ObjType = "��������." + Mid(ObjType, 5)
	elseif left(UCase(ObjType), 4) = "���." then
		ObjType = "����������." + Mid(ObjType, 5) 
	elseif left(UCase(ObjType), 4) = "���." then
		ObjType = "������������" + Mid(ObjType, 5)
	end if
	
	CntArr = GetDefaultStringArr("Text") 
	CntArr(3) = "16" 
	CntArr(4) = "82" 
	CntArr(5) = "200"  
	CntArr(6) = "14"  
	CntArr(9) = "4260"
	CntArr(12) = Ident
	t = SetType(CntArr, ObjType)
	
	DlgBlock = "{""" + Join(CntArr, """,""")
	
	ButtonArray = GetDefaultStringArr("Button") 
	ButtonArray(9) = "4261"
	ButtonArray(0) = "O"
	ButtonArray(11) = "������������(" + Ident +")" 
	ButtonArray(3) = "220" 
	ButtonArray(4) = "82" 
	ButtonArray(5) = "18"  
	ButtonArray(6) = "14" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	ButtonArray(9) = "4262"
	ButtonArray(0) = "X"
	ButtonArray(11) = Ident + " = """"" 
	ButtonArray(3) = "240" 
	ButtonArray(4) = "82" 
	ButtonArray(5) = "18"  
	ButtonArray(6) = "14" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	LabelArr = GetDefaultStringArr("Label") 
	LabelArr(3) = "16" 
	LabelArr(4) = "72" 
	LabelArr(5) = "200"  
	LabelArr(6) = "9" 
	LabelArr(0) = GetSinonimFromID(Ident) + ":" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(LabelArr, """,""")  
	
	TextArr = Split(DlgText, vbCrLf)
	WorkString = Ubound(TextArr) - 1
	TextArr(WorkString) = Left(TextArr(WorkString), Len(TextArr(WorkString)) - 2) + _
		"," + vbCrLf + DlgBlock + "},"
		
	doc.Stream = Join(TextArr, vbCrLf)
	
end sub
	
	
	
		
'---------------------------------------------------------------------------------
Sub DobavitjTZnaFormu()
    
	Set doc = Windows.ActiveWnd.Document.Page(0)
	DlgText = doc.Stream  
	
	Set ModuleDoc = Windows.ActiveWnd.Document.Page(1)  
	ModuleText = ModuleDoc.Text

	VTArray = GetDefaultStringArr("VT") 
	VTArray(9)  = "4260"
	VTArray(12) = Trim(inputBox("������� ������������� ������� ��������"))  
	VTArray(11) = VTArray(12) + "_���������()"
	VTArray(41) = inputBox("������� ����, �� ������� ��������� ��",,"��������") 
	
	DlgBlock = "{""" + Join(VTArray, """,""")
	
	ButtonArray = GetDefaultStringArr("Button") 
	ButtonArray(9) = "4261"
	ButtonArray(0) = "+"
	ButtonArray(11) = VTArray(12) + "_��������������()" 
	ButtonArray(3) = "22" 
	ButtonArray(4) = "26" 
	ButtonArray(5) = "18"  
	ButtonArray(6) = "13" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	ButtonArray(9) = "4262"
	ButtonArray(0) = "-"
	ButtonArray(11) = VTArray(12) + "_�������������()" 
	ButtonArray(3) = "43" 
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	TextArr = Split(DlgText, vbCrLf)
	WorkString = Ubound(TextArr) - 1
	TextArr(WorkString) = Left(TextArr(WorkString), Len(TextArr(WorkString)) - 2) + _
		"," + vbCrLf + DlgBlock + "},"  
		
	TabInd = vbCrLf   + "" 
	
	VTName = VTArray(12)
	ColumnsList = inputBox("������� ����� ������� ����� �������. ��� ������� ����� ���������� ����� ����� '#', " + _
		"����� � �������� �������� � ���� '�����-15-2', ���� ������� ��������� ��������� �� ����� ��������������� ����")  
	VTColumns = Split(ColumnsList, ",")
	
	ChoiceCode = ""
	InvisibleColumns = ""
	Txt = ""  
	for i = 0 to UBound(VTColumns) 
	
		ColumnName = ""
	    ColumnType = ""
		ColumnLength = ""
		ColumnTochn = ""
		
		Pos = instr(VTColumns(i), "#")
		if Pos = 0 then 
			ColumnName = Trim(VTColumns(i)) 
			if Right(ColumnName, 1) = "!" then
				ColumnName = Replace(ColumnName, "!", "")
				InvisibleColumns = InvisibleColumns + "," + ColumnName
			end if
		else
			ColumnName = left(VTColumns(i), Pos - 1)   
			if Right(ColumnName, 1) = "!" then
				ColumnName = Replace(ColumnName, "!", "")
				InvisibleColumns = InvisibleColumns + "," + ColumnName
			end if
			TypeDefine = Mid(VTColumns(i), Pos + 1)  
			if Right(TypeDefine, 1) = "!" then
				TypeDefine = Replace(TypeDefine, "!", "")
				InvisibleColumns = InvisibleColumns + "," + ColumnName
			end if 
			TypeDefineArr = Split(TypeDefine, "-")  
			if Ubound(TypeDefineArr) > - 1 then
				ColumnType = Replace(TypeDefineArr(0), """", "")
				if UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "����������." + Mid(ColumnType, 5)
				elseif UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "��������." + Mid(ColumnType, 5) 
				elseif UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "������������." + Mid(ColumnType, 5) 
				end if
			end if  
			if Ubound(TypeDefineArr) > 0 then
				ColumnLength = TypeDefineArr(1)
			end if 
			if Ubound(TypeDefineArr) > 1 then
				ColumnTochn = TypeDefineArr(2)
			end if
			
		end if
		
		Txt = Txt + TabInd + VTName + ".������������(""" + Trim(ColumnName) + """"
		if ColumnType <> "" then Txt = Txt + ", """ + Trim(ColumnType) + """" 
		if ColumnLength <> "" then Txt = Txt + ", " + Trim(ColumnLength) 
		if ColumnTochn <> "" then Txt = Txt + ", " + Trim(ColumnTochn) 
			
		if i = 0 then
			firstWrd = ""
		else 
			firstWrd = "�����"
		end if
			
		ChoiceCode = ChoiceCode + vbCrLf + vbTab + FirstWrd + "���� " + VTName + ".��������������() = """ + Trim(ColumnName) + """ �����" 
		if ColumnType = "" then
			ChoiceCode = ChoiceCode + vbCrLf
		elseif UCase(ColumnType) = "�����" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "������������� = " + VTName + "." + ColumnName + ";"
			if ColumnLength = "" then 
				ColumnLength = "15"
			end if
			if ColumnTochn = "" then 
				ColumnTochn = "4"
			end if 
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "�����������(�������������, ""������� �����"", " + ColumnLength + ", " + ColumnTochn + ");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = �������������;"
		elseif UCase(ColumnType) = "������" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "�������������� = " + VTName + "." + ColumnName + ";"
			if ColumnLength = "" then 
				ColumnLength = "100"
			end if
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "������������(��������������, ""������� ������"", " + ColumnLength + ");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ��������������;" 
		elseif UCase(ColumnType) = "����" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "������������ = " + VTName + "." + ColumnName + ";"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "����������(������������, ""������� ����"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ������������;" 
		elseif (left(UCase(ColumnType), 12) = "������������") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "��������� = " + VTName + "." + ColumnName + ";"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "������������������(���������, ""�������� ��������"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ���������;"  
		elseif (left(UCase(ColumnType), 10) = "����������") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "��� = �������������(""" + ColumnType + """);"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "���.�������(""�������� �������"", ""��������"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ���.��������������();"  
		elseif (left(UCase(ColumnType), 8) = "��������") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "��� = �������������(""" + ColumnType + """);"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "���.�������(""�������� ��������"", ""������.�����"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ���.���������������();" 
		end if 
		
		Txt = Txt + ");"
	Next   
	if UBound(VTColumns) >= 0 then ChoiceCode = ChoiceCode + vbCrLf + vbTab + "���������;" + vbCrLf
	
	if InvisibleColumns <> "" then
		Txt = Txt + TabInd + VTName + ".����������������(""" + Mid(InvisibleColumns, 2) + """, 0);"	 
	end if  
	
	ModuleArr = Split(ModuleText, vbCrLf)
	ModuleArr(UBound(ModuleArr)) = ModuleArr(UBound(ModuleArr)) + vbCrLf + VbCrLf + Txt  
	
	ProcBlock = VbCrLf + VbCrLf + "//=========================================================" 
	ProcBlock = ProcBlock + vbCrLf + "//���������� ��� ������� ������ �� ������� ��������"
	ProcBlock = ProcBlock + vbCrLf + "��������� " + VTName + "_���������()"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� " + VTName + ".���������������() = 0 �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + VTName + ".�����������();"   
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + ChoiceCode
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + VTName + "_��������������()"
	ProcBlock = ProcBlock + vbCrLf + vbTab + VTName + ".�����������();" 
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + VTName + "_�������������()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� " + VTName + ".����������� > 0 �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + VTName + ".�������������();"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf 
	
	t = AddProcsToModuleBegin(ModuleArr, ProcBlock)
		
	ModuleDoc.Text = Join(ModuleArr, vbCrLf)	
	doc.Stream = Join(TextArr, vbCrLf)
		
End sub
	
'---------------------------------------------------------------------------------
Sub DobavitjSZnaFormu()
    
	Set doc = Windows.ActiveWnd.Document.Page(0)
	DlgText = doc.Stream  
	
	Set ModuleDoc = Windows.ActiveWnd.Document.Page(1)  
	ModuleText = ModuleDoc.Text  

    
	'============== ������ =====================================================
	ListArray = GetDefaultStringArr("List") 
	ListArray(3) = "16" 
	ListArray(4) = "38" 
	ListArray(5) = "200"  
	ListArray(6) = "90" 
	ListArray(9)  = "4260"
	ListArray(12) = Trim(inputBox("������� ������������� ������ ��������"))  
	ListArray(11) = ListArray(12) + "_���������()"
	ListArray(41) = inputBox("������� ����, �� ������� ��������� ������",,"��������")  
	
	ValueTypeDef = inputBox("������� ��� �������� ������") 
	if Instr(ValueTypeDef, ".") > 0 then
		ValueKind = Mid(ValueTypeDef, Instr(ValueTypeDef, ".") + 1)
	else
		ValueKind = ""
	end if 
	
	RamkaArr = GetDefaultStringArr("Frame") 
	RamkaArr(0) = ValueKind
	RamkaArr(3) = "12" 
	RamkaArr(4) = "29" 
	RamkaArr(5) = "231"  
	RamkaArr(6) = "101" 
	
	DlgBlock = "{""" + Join(ListArray, """,""") + "," + vbCrLf + "{""" + Join(RamkaArr, """,""")  
	
	ButtonArray = GetDefaultStringArr("Button") 
	ButtonArray(9) = "4261"
	ButtonArray(0) = "+"
	ButtonArray(11) = ListArray(12) + "_����������������()" 
	ButtonArray(3) = "220" 
	ButtonArray(4) = "38" 
	ButtonArray(5) = "16"  
	ButtonArray(6) = "12" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	ButtonArray(9) = "4262"
	ButtonArray(0) = "-"
	ButtonArray(11) = ListArray(12) + "_���������������()" 
	ButtonArray(4) = "51"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	if (left(UCase(ValueTypeDef), 10) = "����������") And (ValueKind <> "") then
		ButtonArray(9) = "4263"
		ButtonArray(0) = "..."
		ButtonArray(11) = ListArray(12) + "_������������������()" 
		ButtonArray(4) = "66"   
		
		DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	end if
	
	ButtonArray(9) = "4264"
	ButtonArray(0) = "X"
	ButtonArray(11) = ListArray(12) + "_����������()" 
	ButtonArray(4) = "79"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	ButtonArray(9) = "4265"
	ButtonArray(0) = "U"
	ButtonArray(11) = ListArray(12) + "_�������������()" 
	ButtonArray(4) = "95"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	ButtonArray(9) = "4266"
	ButtonArray(0) = "D"
	ButtonArray(11) = ListArray(12) + "_������������()" 
	ButtonArray(4) = "108"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	TextArr = Split(DlgText, vbCrLf)
	WorkString = Ubound(TextArr) - 1
	TextArr(WorkString) = Left(TextArr(WorkString), Len(TextArr(WorkString)) - 2) + _
		"," + vbCrLf + DlgBlock + "},"  
		
	'=====================================================================================  
	
	ListName = ListArray(12)
	
	ProcBlock = VbCrLf + VbCrLf + "//=========================================================" 
	ProcBlock = ProcBlock + vbCrLf + "//���������� ��� ������� ������ �� ������"
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_���������()"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� " + ListName + ".������������() > 0 �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "����������������� = " + ListName + ".����������������(" + ListName + ".�������������());" 
	if Instr(UCase(ValueTypeDef), "����������") = 1 then
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "������������(�����������������.��������������());"
	elseif Instr(UCase(ValueTypeDef), "��������") = 1 then
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "������������(�����������������.���������������());" 
	end if
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_����������������()"
	if ValueTypeDef = "" then
		ProcBlock = ProcBlock + vbCrLf
	elseif UCase(ValueTypeDef) = "�����" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "������������� = 0;"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "�����������(�������������, ""������� �����"", 15, 4);"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".����������������(�������������);"
	elseif UCase(ValueTypeDef) = "������" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "�������������� = """";"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "������������(��������������, ""������� ������"", 100);"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".����������������(��������������);"
	elseif UCase(ValueTypeDef) = "����" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "������������ = ����(0);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "����������(������������, ""������� ����"");"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".����������������(������������);"
	elseif (left(UCase(ValueTypeDef), 12) = "������������") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "��������� = ������������." + ValueKind + ".����������������(1);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + VbTab + "������������������(���������, ""�������� ��������"");"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".����������������(���������);"
	elseif (left(UCase(ValueTypeDef), 10) = "����������") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "��� = �������������(""" + ValueTypeDef + """);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���� ���.�������(""�������� �������"", ""��������"") = 1 �����"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".����������������(���.��������������());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	elseif (left(UCase(ValueTypeDef), 8) = "��������") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "��� = �������������(""" + ValueTypeDef + """);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���� ���.�������(""�������� ��������"") = 1 �����"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".����������������(���.���������������());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	end if 
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_���������������()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� " + ListName + ".�������������() > 0 �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".���������������(" + ListName + ".�������������());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf 
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_����������()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".����������();"  
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf 
    
	if (left(UCase(ValueTypeDef), 10) = "����������") And (ValueKind <> "") then
		ProcBlock = ProcBlock + "//========================================================="
		ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_������������������()" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "�������������(""" + ValueTypeDef + """, ""��������"",,1);"	
		ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf  
		
		ProcBlock = ProcBlock + "//========================================================="
		ProcBlock = ProcBlock + vbCrLf + "��������� ����������������(�����������)" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���� ��������������(�����������) = ""����������"" �����"   
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".����������������(�����������.��������������());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".�������������(" + ListName + ".������������());"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	    ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf
	end if  
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_�������������()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� " + ListName + ".�������������() > 1 �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".����������������(-1, " + ListName + ".�������������());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf  
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "��������� " + ListName + "_������������()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���� (" + ListName + ".�������������() > 0) � (" + ListName + ".�������������() < " +  ListName + ".������������()) �����"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".����������������(1, " + ListName + ".�������������());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "���������;"
	ProcBlock = ProcBlock + vbCrLf + "��������������" + vbCrLf + vbCrLf 

	
	ModuleArr = Split(ModuleText, vbCrLf)
	t = AddProcsToModuleBegin(ModuleArr, ProcBlock)
		
	ModuleDoc.Text = Join(ModuleArr, vbCrLf)	
	doc.Stream = Join(TextArr, vbCrLf)
		
End sub