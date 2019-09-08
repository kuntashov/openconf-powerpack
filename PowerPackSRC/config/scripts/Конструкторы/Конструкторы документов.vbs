 
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

Private Function GetDocumentKind()
	Set t=Plugins("Телепат")
	GetDocumentKind = t.ConvertTemplate("<?""Укажите документ"",Документ>")
End Function	
	
'============================================================================================= 
Sub VybratjDocument()
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
	Kind = GetDocumentKind
	VarName =  "Док" + Kind
	Txt = VarName + " = СоздатьОбъект(""Документ." + Kind + """);" 
    Txt = Txt +  TabInd +  "Если " + VarName +  ".Выбрать(""Выберите документ"") = 1 Тогда"
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЕсли;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt   
	
End Sub

'=============================================================================================
Sub VyborkaDocumentov() 'Цикл выборки документов 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetDocumentKind
	
	VarName =  "Док" + Kind
	Txt = VarName + " = СоздатьОбъект(""Документ." + Kind + """);" 
    Txt  = Txt +  TabInd +  VarName +  ".ВыбратьДокументы();"   
	
	flNeProv = msgbox("Включать в выборку непроведенные документы?", 4) 
	if flNeProv = 6 then
		flPomUd = msgbox("Включать в выборку помеченные на удаление?", 4) 
	else
		flPomUd = 6
		Txt  = Txt +  TabInd +  VarName +  ".УстановитьФильтр(1, 0);"  
	end if
	
	Txt  = Txt +  TabInd +  "Пока " + VarName +  ".ПолучитьДокумент() = 1 Цикл"
	if flPomUd <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "Если " + VarName +  ".ПометкаУдаления() = 1 Тогда"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "Продолжить;"
		Txt  = Txt +  TabInd +  vbTab + "КонецЕсли;" + TabInd
	end if
    Txt = Txt + TabInd + vbTab + "//Код" + TabInd + "КонецЦикла;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub         

'============================================================================================= 
Sub VygruzitjTabChastj()
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
 	VarName = inputBox("Введите имя переменной, содержащей ссылку на документ:") 
 	
 	TypeObj = msgbox("Выгрузить в таблицу значений(нажмите ДА) или в список значений(нажмите НЕТ) ?", 4)
 	if TypeObj = 6 then
		Txt = TabInd + "ТаблицаДокумента = СоздатьОбъект(""ТаблицаЗначений"");"
		ColumnNames = Trim(inputBox("Введите имена выгружаемых колонок через запятую:"))
		
		if ColumnNames = "" then
			Txt = Txt +  TabInd +  VarName + ".ВыгрузитьТабличнуюЧасть(ТаблицаДокумента);"
		else
			Txt = Txt +  TabInd +  VarName + ".ВыгрузитьТабличнуюЧасть(ТаблицаДокумента, """ + _
			Replace(ColumnNames, """", "") + """);"
		end if  
	else
		ColumnName = Replace(Trim(inputBox("Введите имя выгружаемой колонки:")), """", "") 
		ListName = ColumnName + "Список"
		Txt = ListName + " = СоздатьОбъект(""СписокЗначений"");" 
		Txt = Txt +  TabInd +  VarName + ".ВыгрузитьТабличнуюЧасть(" + ListName + ", """ + _
		ColumnName + """);"  
	end if
       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt   
	
End Sub
	
	
 	