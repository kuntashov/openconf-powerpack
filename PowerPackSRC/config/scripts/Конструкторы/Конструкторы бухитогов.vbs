
'=============================================================================================
Function CheckWindow(doc) 
	
	CheckWindow = False
	
	If Windows.ActiveWnd Is Nothing Then
    		MsgBox "Ќет активного окна"
    		Exit Function
	End If   
	
	Set doc = Windows.ActiveWnd.Document
	If doc=docWorkBook Then Set doc=doc.Page(1)  
		
	If doc<>docText Then
    		MsgBox "ќкно не текстовое"
    		Exit Function
	End If   
	
	CheckWindow = True

End Function
    
'=============================================================================================
Sub Zapros() 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf + String( doc.SelStartCol, vbTab)
	        
	VarName   = inputBox("¬ведите им€ переменной бухитогов:",,"»тоги") 
	SubkList  = inputBox("¬ведите имена видов субконто через зап€тую, по которым следует разворачивать итоги")  
	SubkList2 = inputBox("¬ведите имена видов субконто через зап€тую, по которым следует делать отбор") 
	
	Subks = Split(SubkList, ",")
	Txt = VarName + " = —оздатьќбъект(""Ѕухгалтерские»тоги"");" 
	for i = 0 to UBound(Subks) 
		Txt = Txt + TabInd + VarName + ".»спользовать—убконто(¬иды—убконто." + Subks(i) + ");"
	Next
	
	SubksOtbor = Split(SubkList2, ",")	
	for i = 0 to UBound(SubksOtbor)  
		CurrOtbor = inputBox("¬ведите значение отбора дл€ ¬иды субконто " + SubksOtbor(i)) 
		Txt = Txt + TabInd + VarName + ".»спользовать—убконто(¬иды—убконто." + SubksOtbor(i) + ", " + CurrOtbor + ", 2);"
	Next
	
	Params  = inputBox("¬ведите через зап€тую начальную дату запроса, конечную дату запроса и значение счета запроса") 
	Txt = Txt + TabInd + VarName + ".¬ыполнить«апрос(" + Params + ")"
	  
	TextVybEnd = ""
	for i = 0 to UBound(Subks) 
		DopTabInd = ""
		if i > 0 then
			DopTabInd = String(i, vbTab)
		end if
		
		Txt = Txt + TabInd + DopTabInd + VarName + ".¬ыбрать—убконто(¬иды—убконто." + Subks(i) + ");"
		Txt = Txt + TabInd + DopTabInd + "ѕока " + VarName + ".ѕолучить—убконто(¬иды—убконто." + Subks(i) + ") = 1 ÷икл"
		
		TextVybEnd = TabInd + DopTabInd + " онец÷икла;" + TextVybEnd
	Next  
	
	Txt = Txt + vbCrLf + TabInd + TextVybEnd
	
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      