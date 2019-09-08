' IAm
'Первый позволяет быстро инициализировать таблицу значений, введя имена колонок через запятую, второй позволяет по имени таблицы создать новую строку, ищет по тексту строки с методом "НоваяКолонка()" и автоматом прописывает поля. То есть если в модуле был блок
'Табл.НоваяКолонка("Контр");
'Табл.НоваяКолонка("Договор");
'Табл.НоваяКолонка("Сумма");
'то можно вызвать скрипт, ввести имя таблицы(Табл)
'и получить текст в виде:
'Табл.Контр =
'Табл.Договор =
'Табл.Сумма =
'==================================================

'=============================================================================================
Function CheckWindow(doc)

	CheckWindow = False

	If Windows.ActiveWnd Is Nothing Then
    		MsgBox "Нет активного окна"
    		Exit Function
	End If

	Set doc = Windows.ActiveWnd.Document
	If doc=docWorkBook Then Set doc=doc.Page(1)

	If doc<>docText Then
    		MsgBox "Окно не текстовое"
    		Exit Function
	End If

	CheckWindow = True

End Function

'==================================================
Sub NovayaTabliza()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("введите имя таблицы значений:")
	ColumnsList = inputBox("Введите имена колонок через запятую. Тип колонки можно определить после знака '#', " + _
		"Длина и точность задается в виде 'Число-15-2', Если колонка невидимая поставьте на конце восклицательный знак")
	VTColumns = Split(ColumnsList, ",")

	InvisibleColumns = ""
	Txt = VTName + " = СоздатьОбъект(""ТаблицаЗначений"");"
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

		Txt = Txt + ");"
	Next

	if InvisibleColumns <> "" then
		Txt = Txt + TabInd + VTName + ".ВидимостьКолонки(""" + Mid(InvisibleColumns, 2) + """, 0);"
	end if

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'====================================================
Sub NovayaStroka()
' Артур TODO не различает комментарии

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

    VTName = inputBox("введите имя таблицы значений:")
	NewColumnText = UCase(VTName + ".НоваяКолонка(""")
	StrLen = Instr(NewColumnText, """")

    TextD = doc.text
	TextDUpper = UCase(TextD)

	Txt = VTName + ".НоваяСтрока();"

	StartPos = 0
	Pos = instr(TextDUpper, NewColumnText)
    do while Pos > 0
		StartPos = Pos + 10
		Pos2 = instr(Pos + StrLen, TextDUpper, """")
		if Pos2 - Pos - StrLen < 30 then
			Txt = Txt + TabInd + VTName + "." + Mid(TextD, Pos + StrLen, Pos2 - Pos - StrLen) + " = ;"
		end if
		Pos = instr(StartPos, TextDUpper, NewColumnText)
	Loop

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
end sub

'====================================================
Sub PoiskZnacheniya()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("введите имя таблицы значений:")
	ColumnName = inputBox("введите имя колонки поиска")
	valueToFind = inputBox("введите значение поиска, или имя переменной, содержащей значение")

	Txt = "Позиция = 0;"
	Txt = Txt + TabInd + "Если " + VTName + ".НайтиЗначение(" + ValueToFind + ", Позиция, """ + ColumnName + """) = 1 Тогда"
	Txt = Txt + TabInd + vbTab + "//Код"
	Txt = Txt + TabInd + "КонецЕсли;"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'====================================================
Sub PoiskPoDvumKolonkam()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("введите имя таблицы значений:")
	Otbor1 = inputBox("Через запятую введите имя, затем значение отбора первой колонки:")
	Otbor2 = inputBox("То же самое для второй колонки:")

	Pos1 = instr(Otbor1, ",")
	ColumnName1 = Replace(Left(Otbor1, Pos1 - 1), """", "")
	OtborValue1 = Mid(Otbor1, Pos1 + 1)

	Pos2 = instr(Otbor2, ",")
	ColumnName2 = Replace(Left(Otbor2, Pos2 - 1), """", "")
	OtborValue2 = Mid(Otbor2, Pos2 + 1)

	Txt = "Позиция = 0;"
	Txt = Txt + TabInd + "Если " + VTName + ".НайтиЗначение(" + OtborValue1 + ", Позиция, """ + ColumnName1 + """) = 1 Тогда"
	Txt = Txt + TabInd + vbTab + "Для А = Позиция По " + VTName + ".КоличествоСтрок() Цикл"
	Txt = Txt + TabInd + vbTab + vbTab + _
		"Если " + VTName + ".ПолучитьЗначение(А, """ + ColumnName1 + """) <> " + OtborValue1 + " Тогда"
	Txt = Txt + TabInd + vbTab + vbTab + vbTab + "Прервать;"
	Txt = Txt + TabInd + vbTab + vbTab + "КонецЕсли;"
	Txt = Txt + TabInd
	Txt = Txt + TabInd + vbTab + vbTab + _
	    "Если " + VTName + ".ПолучитьЗначение(А, """ + ColumnName2 + """) = " + OtborValue2 + " Тогда"
    Txt = Txt + TabInd + vbTab + vbTab + vbTab + "//Код"
	Txt = Txt + TabInd + vbTab + vbTab + "КонецЕсли;"
	Txt = Txt + TabInd + vbTab + "КонецЦикла;"
	Txt = Txt + TabInd + "КонецЕсли;"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'=====================================================
Sub VygruzkaVSpisokZnacheniy()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("введите имя таблицы значений:")
	ColumnName = inputBox("введите имя выгружаемой колонки:")
	VLName = inputBox("введите имя списка значений:",,ColumnName + "Список")

	Txt = VLName + " = СоздатьОбъект(""СписокЗначений"");"
	Txt = Txt + TabInd + VTName + ".Выгрузить(" + VLName + ",,, """ + Replace(ColumnName, """", "") + """);"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub
