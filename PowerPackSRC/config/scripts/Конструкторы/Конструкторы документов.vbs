 
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

Private Function GetDocumentKind()
	Set t=Plugins("�������")
	GetDocumentKind = t.ConvertTemplate("<?""������� ��������"",��������>")
End Function	
	
'============================================================================================= 
Sub VybratjDocument()
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
	Kind = GetDocumentKind
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""��������." + Kind + """);" 
    Txt = Txt +  TabInd +  "���� " + VarName +  ".�������(""�������� ��������"") = 1 �����"
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "���������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt   
	
End Sub

'=============================================================================================
Sub VyborkaDocumentov() '���� ������� ���������� 

	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab)
	
	Kind = GetDocumentKind
	
	VarName =  "���" + Kind
	Txt = VarName + " = �������������(""��������." + Kind + """);" 
    Txt  = Txt +  TabInd +  VarName +  ".����������������();"   
	
	flNeProv = msgbox("�������� � ������� ������������� ���������?", 4) 
	if flNeProv = 6 then
		flPomUd = msgbox("�������� � ������� ���������� �� ��������?", 4) 
	else
		flPomUd = 6
		Txt  = Txt +  TabInd +  VarName +  ".����������������(1, 0);"  
	end if
	
	Txt  = Txt +  TabInd +  "���� " + VarName +  ".����������������() = 1 ����"
	if flPomUd <> 6 then
		Txt  = Txt +  TabInd +  vbTab + "���� " + VarName +  ".���������������() = 1 �����"
		Txt  = Txt +  TabInd +  vbTab + vbTab + "����������;"
		Txt  = Txt +  TabInd +  vbTab + "���������;" + TabInd
	end if
    Txt = Txt + TabInd + vbTab + "//���" + TabInd + "����������;"       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt
	  
end sub         

'============================================================================================= 
Sub VygruzitjTabChastj()
	
	doc = ""
	if Not CheckWindow(doc) then Exit sub 
	
	TabInd = vbCrLf   + String( doc.SelStartCol, vbTab) 
	
 	VarName = inputBox("������� ��� ����������, ���������� ������ �� ��������:") 
 	
 	TypeObj = msgbox("��������� � ������� ��������(������� ��) ��� � ������ ��������(������� ���) ?", 4)
 	if TypeObj = 6 then
		Txt = TabInd + "���������������� = �������������(""���������������"");"
		ColumnNames = Trim(inputBox("������� ����� ����������� ������� ����� �������:"))
		
		if ColumnNames = "" then
			Txt = Txt +  TabInd +  VarName + ".�����������������������(����������������);"
		else
			Txt = Txt +  TabInd +  VarName + ".�����������������������(����������������, """ + _
			Replace(ColumnNames, """", "") + """);"
		end if  
	else
		ColumnName = Replace(Trim(inputBox("������� ��� ����������� �������:")), """", "") 
		ListName = ColumnName + "������"
		Txt = ListName + " = �������������(""��������������"");" 
		Txt = Txt +  TabInd +  VarName + ".�����������������������(" + ListName + ", """ + _
		ColumnName + """);"  
	end if
       
    doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt   
	
End Sub
	
	
 	