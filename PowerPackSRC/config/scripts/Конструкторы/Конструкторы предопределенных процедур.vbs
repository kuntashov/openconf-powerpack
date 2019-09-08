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
    
'===========================================================================================
Sub PriNachaleVyboraZnacheniya() 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
	Txt = "//====================================" + TabInd + "Процедура ПриНачалеВыбораЗначения(ИдентЭлемента, Флаг)"
	Idents = inputBox("Введите через запятую идентификаторы элементов диалога, чей выбор обрабатывается в особом порядке:")
	IdentsList = Split(Trim(Idents), ",")
	
	for i = 0 To UBound(IdentsList)
		
		if i = 0 then
			Txt = Txt + TabInd + TabInd + vbTab + "Если ИдентЭлемента = """ + IdentsList(i) + """ Тогда"
		else
			Txt = Txt + TabInd  + vbTab + "ИначеЕсли ИдентЭлемента = """ + IdentsList(i) + """ Тогда"   
		end if
		
		Txt = Txt + TabInd + vbTab + vbTab + "Флаг = 0;" + TabInd
	Next
	
	Txt = Txt + TabInd + vbTab + "КонецЕсли;" + TabInd + TabInd + "КонецПроцедуры"
	
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt 
	
End sub