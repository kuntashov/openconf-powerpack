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
    
    
'=============================================================================================
Sub NovayaProvodka()   

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)	
	
	OperName = inputBox("Введите имя объекта 'Операция'",,"Операция")
    DtString = inputBox("Введите через запятую значение счета дебета и последовательно идентификаторы видов субконто:")
    KtString = inputBox("Введите через запятую значение счета кредита и последовательно идентификаторы видов субконто:") 
    
	Txt = OperName + ".НоваяПроводка();"
	
    DtObjs = Split(DtString, ",")
    if UBound(DtObjs) > -1 Then    
		CountName = DtObjs(0)
		if Instr(CountName, """") = 0 then
			if Cint(CountName) > 0 then
				CountName = """" + CountName + """"
			end if
		end if     
		if Instr(CountName, """") = 0 then
    		Txt = Txt + TabInd + OperName + ".Дебет.Счет = " + CountName + ";" 
		else
			Txt = Txt + TabInd + OperName + ".Дебет.Счет = СчетПоКоду(" + CountName + ");"
		end if
        
        for i = 1 to UBound(DtObjs) 
			SubkString = OperName + ".Дебет." + DtObjs(i) + " = "
        	SubkValue = inputBox(SubkString)
            Txt = Txt + TabInd + SubkString + SubkValue + ";"
        Next
        
    end if
    
    KtObjs = Split(KtString, ",")
    if UBound(KtObjs) > -1 Then    
		CountName = KtObjs(0)
		if Instr(CountName, """") = 0 then
			if Cint(CountName) > 0 then
				CountName = """" + CountName + """"
			end if
		end if
    	if Instr(CountName, """") = 0 then
    		Txt = Txt + TabInd + OperName + ".Кредит.Счет = " + CountName + ";" 
		else
			Txt = Txt + TabInd + OperName + ".Кредит.Счет = СчетПоКоду(" + CountName + ");"
		end if
        
        for i = 1 to UBound(KtObjs) 
			SubkString = OperName + ".Кредит." + KtObjs(i) + " = "
        	SubkValue = inputBox(SubkString)
            Txt = Txt + TabInd + SubkString + SubkValue + ";"
        Next
        
    end if  
    
	Txt = Txt + TabInd + OperName + ".Сумма = " + InputBox(OperName + ".Сумма = ") + ";" 
	Txt = Txt + TabInd + OperName + ".Количество = " + InputBox(OperName + ".Количество = ") + ";"
    Txt = Txt + TabInd + OperName + ".СодержаниеПроводки = """";"  
    
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
    
End sub