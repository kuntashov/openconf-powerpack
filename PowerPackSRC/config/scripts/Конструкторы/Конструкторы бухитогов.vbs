
'=============================================================================================
Function CheckWindow(doc) 
	
	CheckWindow = False
	
	If Windows.ActiveWnd Is Nothing Then
    		MsgBox "��� ��������� ����"
    		Exit Function
	End If   
	
	Set doc = Windows.ActiveWnd.Document
	If doc=docWorkBook Then Set doc=doc.Page(1)  
		
	If doc<>docText Then
    		MsgBox "���� �� ���������"
    		Exit Function
	End If   
	
	CheckWindow = True

End Function
    
'=============================================================================================
Sub Zapros() 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf + String( doc.SelStartCol, vbTab)
	        
	VarName   = inputBox("������� ��� ���������� ���������:",,"�����") 
	SubkList  = inputBox("������� ����� ����� �������� ����� �������, �� ������� ������� ������������� �����")  
	SubkList2 = inputBox("������� ����� ����� �������� ����� �������, �� ������� ������� ������ �����") 
	
	Subks = Split(SubkList, ",")
	Txt = VarName + " = �������������(""������������������"");" 
	for i = 0 to UBound(Subks) 
		Txt = Txt + TabInd + VarName + ".��������������������(������������." + Subks(i) + ");"
	Next
	
	SubksOtbor = Split(SubkList2, ",")	
	for i = 0 to UBound(SubksOtbor)  
		CurrOtbor = inputBox("������� �������� ������ ��� ���� �������� " + SubksOtbor(i)) 
		Txt = Txt + TabInd + VarName + ".��������������������(������������." + SubksOtbor(i) + ", " + CurrOtbor + ", 2);"
	Next
	
	Params  = inputBox("������� ����� ������� ��������� ���� �������, �������� ���� ������� � �������� ����� �������") 
	Txt = Txt + TabInd + VarName + ".���������������(" + Params + ")"
	  
	TextVybEnd = ""
	for i = 0 to UBound(Subks) 
		DopTabInd = ""
		if i > 0 then
			DopTabInd = String(i, vbTab)
		end if
		
		Txt = Txt + TabInd + DopTabInd + VarName + ".���������������(������������." + Subks(i) + ");"
		Txt = Txt + TabInd + DopTabInd + "���� " + VarName + ".����������������(������������." + Subks(i) + ") = 1 ����"
		
		TextVybEnd = TabInd + DopTabInd + "����������;" + TextVybEnd
	Next  
	
	Txt = Txt + vbCrLf + TabInd + TextVybEnd
	
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      