' IAm
'������ ��������� ������ ���������������� ������� ��������, ����� ����� ������� ����� �������, ������ ��������� �� ����� ������� ������� ����� ������, ���� �� ������ ������ � ������� "������������()" � ��������� ����������� ����. �� ���� ���� � ������ ��� ����
'����.������������("�����");
'����.������������("�������");
'����.������������("�����");
'�� ����� ������� ������, ������ ��� �������(����)
'� �������� ����� � ����:
'����.����� =
'����.������� =
'����.����� =
'==================================================

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

'==================================================
Sub NovayaTabliza()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("������� ��� ������� ��������:")
	ColumnsList = inputBox("������� ����� ������� ����� �������. ��� ������� ����� ���������� ����� ����� '#', " + _
		"����� � �������� �������� � ���� '�����-15-2', ���� ������� ��������� ��������� �� ����� ��������������� ����")
	VTColumns = Split(ColumnsList, ",")

	InvisibleColumns = ""
	Txt = VTName + " = �������������(""���������������"");"
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
				if UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "����������." + Mid(ColumnType, 5)
				elseif UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "��������." + Mid(ColumnType, 5)
				elseif UCase(left(ColumnType, 4)) = "���." then
					ColumnType = "������������." + Mid(ColumnType, 5)
				end if
			end if
			if Ubound(TypeDefineArr) > 0 then
				ColumnLength = TypeDefineArr(1)
			end if
			if Ubound(TypeDefineArr) > 1 then
				ColumnTochn = TypeDefineArr(2)
			end if

		end if

		Txt = Txt + TabInd + VTName + ".������������(""" + Trim(ColumnName) + """"
		if ColumnType <> "" then Txt = Txt + ", """ + Trim(ColumnType) + """"
		if ColumnLength <> "" then Txt = Txt + ", " + Trim(ColumnLength)
		if ColumnTochn <> "" then Txt = Txt + ", " + Trim(ColumnTochn)

		Txt = Txt + ");"
	Next

	if InvisibleColumns <> "" then
		Txt = Txt + TabInd + VTName + ".����������������(""" + Mid(InvisibleColumns, 2) + """, 0);"
	end if

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'====================================================
Sub NovayaStroka()
' ����� TODO �� ��������� �����������

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

    VTName = inputBox("������� ��� ������� ��������:")
	NewColumnText = UCase(VTName + ".������������(""")
	StrLen = Instr(NewColumnText, """")

    TextD = doc.text
	TextDUpper = UCase(TextD)

	Txt = VTName + ".�����������();"

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

	VTName = inputBox("������� ��� ������� ��������:")
	ColumnName = inputBox("������� ��� ������� ������")
	valueToFind = inputBox("������� �������� ������, ��� ��� ����������, ���������� ��������")

	Txt = "������� = 0;"
	Txt = Txt + TabInd + "���� " + VTName + ".�������������(" + ValueToFind + ", �������, """ + ColumnName + """) = 1 �����"
	Txt = Txt + TabInd + vbTab + "//���"
	Txt = Txt + TabInd + "���������;"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'====================================================
Sub PoiskPoDvumKolonkam()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("������� ��� ������� ��������:")
	Otbor1 = inputBox("����� ������� ������� ���, ����� �������� ������ ������ �������:")
	Otbor2 = inputBox("�� �� ����� ��� ������ �������:")

	Pos1 = instr(Otbor1, ",")
	ColumnName1 = Replace(Left(Otbor1, Pos1 - 1), """", "")
	OtborValue1 = Mid(Otbor1, Pos1 + 1)

	Pos2 = instr(Otbor2, ",")
	ColumnName2 = Replace(Left(Otbor2, Pos2 - 1), """", "")
	OtborValue2 = Mid(Otbor2, Pos2 + 1)

	Txt = "������� = 0;"
	Txt = Txt + TabInd + "���� " + VTName + ".�������������(" + OtborValue1 + ", �������, """ + ColumnName1 + """) = 1 �����"
	Txt = Txt + TabInd + vbTab + "��� � = ������� �� " + VTName + ".���������������() ����"
	Txt = Txt + TabInd + vbTab + vbTab + _
		"���� " + VTName + ".����������������(�, """ + ColumnName1 + """) <> " + OtborValue1 + " �����"
	Txt = Txt + TabInd + vbTab + vbTab + vbTab + "��������;"
	Txt = Txt + TabInd + vbTab + vbTab + "���������;"
	Txt = Txt + TabInd
	Txt = Txt + TabInd + vbTab + vbTab + _
	    "���� " + VTName + ".����������������(�, """ + ColumnName2 + """) = " + OtborValue2 + " �����"
    Txt = Txt + TabInd + vbTab + vbTab + vbTab + "//���"
	Txt = Txt + TabInd + vbTab + vbTab + "���������;"
	Txt = Txt + TabInd + vbTab + "����������;"
	Txt = Txt + TabInd + "���������;"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub

'=====================================================
Sub VygruzkaVSpisokZnacheniy()

	doc = ""
	if Not CheckWindow(doc) then Exit sub

	TabInd = vbCrLf   + String(doc.SelStartCol, vbTab)

	VTName = inputBox("������� ��� ������� ��������:")
	ColumnName = inputBox("������� ��� ����������� �������:")
	VLName = inputBox("������� ��� ������ ��������:",,ColumnName + "������")

	Txt = VLName + " = �������������(""��������������"");"
	Txt = Txt + TabInd + VTName + ".���������(" + VLName + ",,, """ + Replace(ColumnName, """", "") + """);"

	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub
