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
    
'===========================================================================================
Sub PriNachaleVyboraZnacheniya() 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
	Txt = "//====================================" + TabInd + "��������� �����������������������(�������������, ����)"
	Idents = inputBox("������� ����� ������� �������������� ��������� �������, ��� ����� �������������� � ������ �������:")
	IdentsList = Split(Trim(Idents), ",")
	
	for i = 0 To UBound(IdentsList)
		
		if i = 0 then
			Txt = Txt + TabInd + TabInd + vbTab + "���� ������������� = """ + IdentsList(i) + """ �����"
		else
			Txt = Txt + TabInd  + vbTab + "��������� ������������� = """ + IdentsList(i) + """ �����"   
		end if
		
		Txt = Txt + TabInd + vbTab + vbTab + "���� = 0;" + TabInd
	Next
	
	Txt = Txt + TabInd + vbTab + "���������;" + TabInd + TabInd + "��������������"
	
	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt 
	
End sub