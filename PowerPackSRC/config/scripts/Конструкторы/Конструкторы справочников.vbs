
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
	
Private Function GetRefKind()
	Set t=Plugins("�������")
	GetRefKind = t.ConvertTemplate("<?""������� ����������"",����������>")
End Function	

'=============================================================================================
Sub VyborElementa() '�������� ����������������� ������� ��� ������ �������� 
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""����������." + Kind + """);" 
    Txt  = Txt +  TabInd +  "���� " + VarName +  ".�������(""�������� �������"", ""�����������"") = 1 �����"
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "���������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      

'=============================================================================================
Sub VyborkaElementov() '���� ������� ��������� 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""����������." + Kind + """);" 
    Txt  = Txt +  TabInd +  VarName +  ".���������������();"
	Txt  = Txt +  TabInd +  "���� " + VarName +  ".���������������() = 1 ����"
	if msgbox("�������� � ������� ������?", 4) <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "���� " + VarName +  ".���������() = 1 �����"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "����������;"
		Txt  = Txt +  TabInd +  vbTab + "���������;" + TabInd
	end if
	if msgbox("�������� � ������� ���������� �� �������� ��������?", 4) <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "���� " + VarName +  ".���������������() = 1 �����"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "����������;"
		Txt  = Txt +  TabInd +  vbTab + "���������;" + TabInd
	end if
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "����������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub      

'=============================================================================================
Sub NaitiElement()   '����� ��������   

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	ValueToFind =  inputBox("������� ��� ����������, ���������� �������� ������:")     
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""����������." + Kind + """);" 
    Txt  = Txt +  TabInd +  "���� " + VarName +  ".������������(" + ValueToFind + ") = 1 �����"
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "���������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub
 
'=============================================================================================
Sub PoiskPoNaimenovaniyu()   '����� �� ������������  

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	SprName =  inputBox("������� ������������ ��� ������ � ��������, ���� ����������, ���������� ��������� �������� ������������")     
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""����������." + Kind + """);" 
    Txt = Txt +  TabInd +  "���� " + VarName +  ".�������������������(" + SprName + ", 0, 1) = 1 �����"
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "���������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub  

'=============================================================================================
Sub PoiskPoKodu()   '����� �� ���� 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetRefKind
	SprName =  inputBox("������� ��� ��� ������ � ��������, ���� ����������, ���������� �������� ����")     
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""����������." + Kind + """);" 
    Txt = Txt +  TabInd +  "���� " + VarName +  ".�����������(" + SprName + ", 0) = 1 �����"
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "���������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub