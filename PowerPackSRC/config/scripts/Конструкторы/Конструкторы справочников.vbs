
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
	
Private Function GetRefKind()
	Set t=Plugins("Телепат")
	GetRefKind = t.ConvertTemplate("<?""Укажите справочник"",Справочник>")
End Function	

'=============================================================================================
Sub VyborElementa() 'Открытие пользовательского диалога для выбора элемента 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	VarName =  "Спр" + Kind
	Txt = VarName + " = СоздатьОбъект(""Справочник." + Kind + """);" 
    Txt  = Txt +  TabInd +  "Если " + VarName +  ".Выбрать(""Выберите элемент"", ""ФормаСписка"") = 1 Тогда"
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЕсли;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      

'=============================================================================================
Sub VyborkaElementov() 'Цикл выборки элементов 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	VarName =  "Спр" + Kind
	Txt = VarName + " = СоздатьОбъект(""Справочник." + Kind + """);" 
    Txt  = Txt +  TabInd +  VarName +  ".ВыбратьЭлементы();"
	Txt  = Txt +  TabInd +  "Пока " + VarName +  ".ПолучитьЭлемент() = 1 Цикл"
	if msgbox("Включать в выборку группы?", 4) <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "Если " + VarName +  ".ЭтоГруппа() = 1 Тогда"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "Продолжить;"
		Txt  = Txt +  TabInd +  vbTab + "КонецЕсли;" + TabInd
	end if
	if msgbox("Включать в выборку Помеченные на удаление элементы?", 4) <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "Если " + VarName +  ".ПометкаУдаления() = 1 Тогда"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "Продолжить;"
		Txt  = Txt +  TabInd +  vbTab + "КонецЕсли;" + TabInd
	end if
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЦикла;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      

'=============================================================================================
Sub NaitiElement()   'Поиск элемента   

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	ValueToFind =  inputBox("Введите имя переменной, содержащей значение поиска:")     
	VarName =  "Спр" + Kind
	Txt = VarName + " = СоздатьОбъект(""Справочник." + Kind + """);" 
    Txt  = Txt +  TabInd +  "Если " + VarName +  ".НайтиЭлемент(" + ValueToFind + ") = 1 Тогда"
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЕсли;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub
 
'=============================================================================================
Sub PoiskPoNaimenovaniyu()   'Поиск по наименованию  

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	SprName =  inputBox("Введите наименование для поиска в кавычках, либо переменную, содержащую строковое значение наименования")     
	VarName =  "Спр" + Kind
	Txt = VarName + " = СоздатьОбъект(""Справочник." + Kind + """);" 
    Txt = Txt +  TabInd +  "Если " + VarName +  ".НайтиПоНаименованию(" + SprName + ", 0, 1) = 1 Тогда"
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЕсли;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub  

'=============================================================================================
Sub PoiskPoKodu()   'Поиск по коду 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	SprName =  inputBox("Введите код для поиска в кавычках, либо переменную, содержащую значение кода")     
	VarName =  "Спр" + Kind
	Txt = VarName + " = СоздатьОбъект(""Справочник." + Kind + """);" 
    Txt = Txt +  TabInd +  "Если " + VarName +  ".НайтиПоКоду(" + SprName + ", 0) = 1 Тогда"
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЕсли;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub