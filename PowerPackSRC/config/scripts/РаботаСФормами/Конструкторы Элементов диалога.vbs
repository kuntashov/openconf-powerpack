'-------------------------------------------------------------------------------
Function GetDefaultStringArr(CtrlType)
    Select Case CtrlType
    Case "Button"
        GetDefaultString = "{""Кнопка"",""BUTTON"",""1342177291"",""8"",""22"",""41"",""13"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "CheckBox"
        GetDefaultString = "{""Текст"",""CHECKBOX"",""1342177283"",""7"",""39"",""43"",""13"",""0"",""0"",""4159"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "Switch"
        GetDefaultString = "{""Текст"",""RADIO"",""1342177289"",""7"",""54"",""48"",""13"",""0"",""0"",""4158"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "List"
        GetDefaultString = "{"""",""LISTBOX"",""1352663297"",""22"",""40"",""220"",""90"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "Combo"
        GetDefaultString = "{"""",""COMBOBOX"",""1352663107"",""7"",""87"",""55"",""15"",""0"",""0"",""4266"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "Frame"
        GetDefaultString = "{""Рамка группы"",""1CGROUPBOX"",""1342177287"",""7"",""105"",""58"",""16"",""0"",""0"",""4155"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "Label"
        GetDefaultString = "{""Текст"",""STATIC"",""1342177280"",""7"",""126"",""63"",""13"",""0"",""0"",""4154"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""192"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""400"",""0"",""0"",""0"",""204"",""1"",""2"",""1"",""34"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "Text"
        GetDefaultString = "{"""",""1CEDIT"",""1350565888"",""7"",""144"",""66"",""13"",""0"",""0"",""4153"","""","""","""",""-1"",""S"",""10"",""0"",""0"",""0"",""0"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "VT"
        GetDefaultString = "{"""",""TABLE"",""1352663040"",""22"",""40"",""220"",""90"",""0"",""0"",""4260"","""","""","""",""0"",""U"",""0"",""0"",""0"",""0"",""16777216"","""","""","""",""0"",""-11"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""0"",""MS Sans Serif"",""-1"",""-1"",""0"",""Основной"",""{""""0"""",""""0""""}""}"
    Case "TableColumn"
        GetDefaultString = "{""4"",""Новая"",""70"",""STATIC"",""4158"","""","""","""",""0"",""2"",""10"",""0"",""0"",""0"",""2"","""",""0"",""268451840"","""","""","""",""0""},"
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
		if (instr(NextLine, "ПРОЦЕДУРА") > 0 Or instr(NextLine, "ФУНКЦИЯ") > 0 ) And _
			instr(NextLine, " ДАЛЕЕ") = 0 then
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
	
	If Instr(ObjTypeMain, "ЧИСЛО") = 1 then
		Params = Split(ObjTypeMain, "-")
		cntType = "N"
		if UBound(Params) > 0 then
			cntLength = Params(1)
		end if
		if UBound(Params) > 1 then
			cntTochn = Params(2)
		end if 
		
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "СТРОКА") = 1 then
		Params = Split(ObjTypeMain, "-")
		cntType = "S"
		if UBound(Params) > 0 then
			cntLength = Params(1)
		end if
		
	elseIf Instr(ObjTypeMain, "ДАТА") = 1 then

		cntType = "D"
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "СЧЕТ") = 1 then

		cntType = "T"
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "СПРАВОЧНИК") = 1 then
	
	    cntType = "B" 
		NumberMD = 1  
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "ДОКУМЕНТ") = 1 then
	
	    cntType = "O" 
		NumberMD = 2
		CntView = "BMASKED"
		
	elseIf Instr(ObjTypeMain, "ПЕРЕЧИСЛЕНИЕ") = 1 then
	
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
	
	Ident = inputBox("Введите идентификатор поля ввода")
	ObjType = inputBox("Введите тип значения. Справочники можно вводить в виде Спр.Контрагенты, " + _
		"Документы - Док.Счет, Перечисления - Пер.ВидыНоменклатуры, Числа - Число-15-2, Строки - Строка-100")
		
	if left(UCase(ObjType), 4) = "ДОК." then
		ObjType = "Документ." + Mid(ObjType, 5)
	elseif left(UCase(ObjType), 4) = "СПР." then
		ObjType = "Справочник." + Mid(ObjType, 5) 
	elseif left(UCase(ObjType), 4) = "ПЕР." then
		ObjType = "Перечисление" + Mid(ObjType, 5)
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
	ButtonArray(11) = "ОткрытьФорму(" + Ident +")" 
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
	VTArray(12) = Trim(inputBox("Введите идентификатор таблицы значений"))  
	VTArray(11) = VTArray(12) + "_ПриВыборе()"
	VTArray(41) = inputBox("Введите слой, на который поместить ТЗ",,"Основной") 
	
	DlgBlock = "{""" + Join(VTArray, """,""")
	
	ButtonArray = GetDefaultStringArr("Button") 
	ButtonArray(9) = "4261"
	ButtonArray(0) = "+"
	ButtonArray(11) = VTArray(12) + "_ДобавитьСтроку()" 
	ButtonArray(3) = "22" 
	ButtonArray(4) = "26" 
	ButtonArray(5) = "18"  
	ButtonArray(6) = "13" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	ButtonArray(9) = "4262"
	ButtonArray(0) = "-"
	ButtonArray(11) = VTArray(12) + "_УдалитьСтроку()" 
	ButtonArray(3) = "43" 
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	TextArr = Split(DlgText, vbCrLf)
	WorkString = Ubound(TextArr) - 1
	TextArr(WorkString) = Left(TextArr(WorkString), Len(TextArr(WorkString)) - 2) + _
		"," + vbCrLf + DlgBlock + "},"  
		
	TabInd = vbCrLf   + "" 
	
	VTName = VTArray(12)
	ColumnsList = inputBox("Введите имена колонок через запятую. Тип колонки можно определить после знака '#', " + _
		"Длина и точность задается в виде 'Число-15-2', Если колонка невидимая поставьте на конце восклицательный знак")  
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
				if UCase(left(ColumnType, 4)) = "СПР." then
					ColumnType = "Справочник." + Mid(ColumnType, 5)
				elseif UCase(left(ColumnType, 4)) = "ДОК." then
					ColumnType = "Документ." + Mid(ColumnType, 5) 
				elseif UCase(left(ColumnType, 4)) = "ПЕР." then
					ColumnType = "Перечисление." + Mid(ColumnType, 5) 
				end if
			end if  
			if Ubound(TypeDefineArr) > 0 then
				ColumnLength = TypeDefineArr(1)
			end if 
			if Ubound(TypeDefineArr) > 1 then
				ColumnTochn = TypeDefineArr(2)
			end if
			
		end if
		
		Txt = Txt + TabInd + VTName + ".НоваяКолонка(""" + Trim(ColumnName) + """"
		if ColumnType <> "" then Txt = Txt + ", """ + Trim(ColumnType) + """" 
		if ColumnLength <> "" then Txt = Txt + ", " + Trim(ColumnLength) 
		if ColumnTochn <> "" then Txt = Txt + ", " + Trim(ColumnTochn) 
			
		if i = 0 then
			firstWrd = ""
		else 
			firstWrd = "Иначе"
		end if
			
		ChoiceCode = ChoiceCode + vbCrLf + vbTab + FirstWrd + "Если " + VTName + ".ТекущаяКолонка() = """ + Trim(ColumnName) + """ Тогда" 
		if ColumnType = "" then
			ChoiceCode = ChoiceCode + vbCrLf
		elseif UCase(ColumnType) = "ЧИСЛО" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ПараметрЧисло = " + VTName + "." + ColumnName + ";"
			if ColumnLength = "" then 
				ColumnLength = "15"
			end if
			if ColumnTochn = "" then 
				ColumnTochn = "4"
			end if 
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ВвестиЧисло(ПараметрЧисло, ""Введите число"", " + ColumnLength + ", " + ColumnTochn + ");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ПараметрЧисло;"
		elseif UCase(ColumnType) = "СТРОКА" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ПараметрСтрока = " + VTName + "." + ColumnName + ";"
			if ColumnLength = "" then 
				ColumnLength = "100"
			end if
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ВвестиСтроку(ПараметрСтрока, ""Введите строку"", " + ColumnLength + ");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ПараметрСтрока;" 
		elseif UCase(ColumnType) = "ДАТА" then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ПараметрДата = " + VTName + "." + ColumnName + ";"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ВвестиДату(ПараметрДата, ""Введите дату"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ПараметрДата;" 
		elseif (left(UCase(ColumnType), 12) = "ПЕРЕЧИСЛЕНИЕ") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ПараметрП = " + VTName + "." + ColumnName + ";"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "ВвестиПеречисление(ПараметрП, ""Выберите значение"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = ПараметрП;"  
		elseif (left(UCase(ColumnType), 10) = "СПРАВОЧНИК") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "Спр = СоздатьОбъект(""" + ColumnType + """);"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "Спр.Выбрать(""Выберите элемент"", ""Основная"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = Спр.ТекущийЭлемент();"  
		elseif (left(UCase(ColumnType), 8) = "ДОКУМЕНТ") And (Instr(ColumnType, ".") > 0) then
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "Док = СоздатьОбъект(""" + ColumnType + """);"
			ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + "Док.Выбрать(""Выберите документ"", ""Журнал.Общий"");"
		    ChoiceCode = ChoiceCode + vbCrLf + vbTab + VbTab + VTName + "." + ColumnName + " = Док.ТекущийДокумент();" 
		end if 
		
		Txt = Txt + ");"
	Next   
	if UBound(VTColumns) >= 0 then ChoiceCode = ChoiceCode + vbCrLf + vbTab + "КонецЕсли;" + vbCrLf
	
	if InvisibleColumns <> "" then
		Txt = Txt + TabInd + VTName + ".ВидимостьКолонки(""" + Mid(InvisibleColumns, 2) + """, 0);"	 
	end if  
	
	ModuleArr = Split(ModuleText, vbCrLf)
	ModuleArr(UBound(ModuleArr)) = ModuleArr(UBound(ModuleArr)) + vbCrLf + VbCrLf + Txt  
	
	ProcBlock = VbCrLf + VbCrLf + "//=========================================================" 
	ProcBlock = ProcBlock + vbCrLf + "//Вызывается при двойном щелчке по таблице значений"
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + VTName + "_ПриВыборе()"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если " + VTName + ".КоличествоСтрок() = 0 Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + VTName + ".НоваяСтрока();"   
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + ChoiceCode
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + VTName + "_ДобавитьСтроку()"
	ProcBlock = ProcBlock + vbCrLf + vbTab + VTName + ".НоваяСтрока();" 
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + VTName + "_УдалитьСтроку()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если " + VTName + ".НомерСтроки > 0 Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + VTName + ".УдалитьСтроку();"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf 
	
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

    
	'============== ДИАЛОГ =====================================================
	ListArray = GetDefaultStringArr("List") 
	ListArray(3) = "16" 
	ListArray(4) = "38" 
	ListArray(5) = "200"  
	ListArray(6) = "90" 
	ListArray(9)  = "4260"
	ListArray(12) = Trim(inputBox("Введите идентификатор списка значений"))  
	ListArray(11) = ListArray(12) + "_ПриВыборе()"
	ListArray(41) = inputBox("Введите слой, на который поместить список",,"Основной")  
	
	ValueTypeDef = inputBox("Введите тип значений списка") 
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
	ButtonArray(11) = ListArray(12) + "_ДобавитьЗначение()" 
	ButtonArray(3) = "220" 
	ButtonArray(4) = "38" 
	ButtonArray(5) = "16"  
	ButtonArray(6) = "12" 
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	
	ButtonArray(9) = "4262"
	ButtonArray(0) = "-"
	ButtonArray(11) = ListArray(12) + "_УдалитьЗначение()" 
	ButtonArray(4) = "51"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	if (left(UCase(ValueTypeDef), 10) = "СПРАВОЧНИК") And (ValueKind <> "") then
		ButtonArray(9) = "4263"
		ButtonArray(0) = "..."
		ButtonArray(11) = ListArray(12) + "_МножественныйВыбор()" 
		ButtonArray(4) = "66"   
		
		DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""")  
	end if
	
	ButtonArray(9) = "4264"
	ButtonArray(0) = "X"
	ButtonArray(11) = ListArray(12) + "_УдалитьВсе()" 
	ButtonArray(4) = "79"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	ButtonArray(9) = "4265"
	ButtonArray(0) = "U"
	ButtonArray(11) = ListArray(12) + "_ЗначениеВверх()" 
	ButtonArray(4) = "95"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	ButtonArray(9) = "4266"
	ButtonArray(0) = "D"
	ButtonArray(11) = ListArray(12) + "_ЗначениеВниз()" 
	ButtonArray(4) = "108"   
	
	DlgBlock = DlgBlock + "," + vbCrLf + "{""" + Join(ButtonArray, """,""") 
	
	TextArr = Split(DlgText, vbCrLf)
	WorkString = Ubound(TextArr) - 1
	TextArr(WorkString) = Left(TextArr(WorkString), Len(TextArr(WorkString)) - 2) + _
		"," + vbCrLf + DlgBlock + "},"  
		
	'=====================================================================================  
	
	ListName = ListArray(12)
	
	ProcBlock = VbCrLf + VbCrLf + "//=========================================================" 
	ProcBlock = ProcBlock + vbCrLf + "//Вызывается при двойном щелчке по списку"
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_ПриВыборе()"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если " + ListName + ".РазмерСписка() > 0 Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "ВыбранноеЗначение = " + ListName + ".ПолучитьЗначение(" + ListName + ".ТекущаяСтрока());" 
	if Instr(UCase(ValueTypeDef), "СПРАВОЧНИК") = 1 then
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "ОткрытьФорму(ВыбранноеЗначение.ТекущийЭлемент());"
	elseif Instr(UCase(ValueTypeDef), "ДОКУМЕНТ") = 1 then
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + "ОткрытьФорму(ВыбранноеЗначение.ТекущийДокумент());" 
	end if
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_ДобавитьЗначение()"
	if ValueTypeDef = "" then
		ProcBlock = ProcBlock + vbCrLf
	elseif UCase(ValueTypeDef) = "ЧИСЛО" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ПараметрЧисло = 0;"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ВвестиЧисло(ПараметрЧисло, ""Введите число"", 15, 4);"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".ДобавитьЗначение(ПараметрЧисло);"
	elseif UCase(ValueTypeDef) = "СТРОКА" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ПараметрСтрока = """";"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ВвестиСтроку(ПараметрСтрока, ""Введите строку"", 100);"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".ДобавитьЗначение(ПараметрСтрока);"
	elseif UCase(ValueTypeDef) = "ДАТА" then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ПараметрДата = Дата(0);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ВвестиДату(ПараметрДата, ""Введите дату"");"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".ДобавитьЗначение(ПараметрДата);"
	elseif (left(UCase(ValueTypeDef), 12) = "ПЕРЕЧИСЛЕНИЕ") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ПараметрП = Перечисление." + ValueKind + ".ЗначениеПоНомеру(1);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + VbTab + "ВвестиПеречисление(ПараметрП, ""Выберите значение"");"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".ДобавитьЗначение(ПараметрП);"
	elseif (left(UCase(ValueTypeDef), 10) = "СПРАВОЧНИК") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "Спр = СоздатьОбъект(""" + ValueTypeDef + """);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "Если Спр.Выбрать(""Выберите элемент"", ""Основная"") = 1 Тогда"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".ДобавитьЗначение(Спр.ТекущийЭлемент());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	elseif (left(UCase(ValueTypeDef), 8) = "ДОКУМЕНТ") And (ValueKind <> "") then
		ProcBlock = ProcBlock + vbCrLf + vbTab + "Док = СоздатьОбъект(""" + ValueTypeDef + """);"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "Если Док.Выбрать(""Выберите документ"") = 1 Тогда"
	    ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".ДобавитьЗначение(Док.ТекущийДокумент());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	end if 
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_УдалитьЗначение()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если " + ListName + ".ТекущаяСтрока() > 0 Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".УдалитьЗначение(" + ListName + ".ТекущаяСтрока());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf 
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_УдалитьВсе()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + ListName + ".УдалитьВсе();"  
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf 
    
	if (left(UCase(ValueTypeDef), 10) = "СПРАВОЧНИК") And (ValueKind <> "") then
		ProcBlock = ProcBlock + "//========================================================="
		ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_МножественныйВыбор()" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "ОткрытьПодбор(""" + ValueTypeDef + """, ""Основная"",,1);"	
		ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf  
		
		ProcBlock = ProcBlock + "//========================================================="
		ProcBlock = ProcBlock + vbCrLf + "Процедура ОбработкаПодбора(ВыбЗначение)" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + "Если ТипЗначенияСтр(ВыбЗначение) = ""Справочник"" Тогда"   
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".ДобавитьЗначение(ВыбЗначение.ТекущийЭлемент());" 
		ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".ТекущаяСтрока(" + ListName + ".РазмерСписка());"
		ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	    ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf
	end if  
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_ЗначениеВверх()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если " + ListName + ".ТекущаяСтрока() > 1 Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".СдвинутьЗначение(-1, " + ListName + ".ТекущаяСтрока());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf  
	
	ProcBlock = ProcBlock + "//========================================================="
	ProcBlock = ProcBlock + vbCrLf + "Процедура " + ListName + "_ЗначениеВниз()" 
	ProcBlock = ProcBlock + vbCrLf + vbTab + "Если (" + ListName + ".ТекущаяСтрока() > 0) И (" + ListName + ".ТекущаяСтрока() < " +  ListName + ".РазмерСписка()) Тогда"
	ProcBlock = ProcBlock + vbCrLf + vbTab + vbTab + ListName + ".СдвинутьЗначение(1, " + ListName + ".ТекущаяСтрока());"  
	ProcBlock = ProcBlock + vbCrLf + vbTab + "КонецЕсли;"
	ProcBlock = ProcBlock + vbCrLf + "КонецПроцедуры" + vbCrLf + vbCrLf 

	
	ModuleArr = Split(ModuleText, vbCrLf)
	t = AddProcsToModuleBegin(ModuleArr, ProcBlock)
		
	ModuleDoc.Text = Join(ModuleArr, vbCrLf)	
	doc.Stream = Join(TextArr, vbCrLf)
		
End sub