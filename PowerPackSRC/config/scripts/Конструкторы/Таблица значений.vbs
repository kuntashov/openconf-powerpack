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
Sub NovayaTabliza()
  If Windows.ActiveWnd Is Nothing Then
     MsgBox "��� ��������� ����"
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     MsgBox "���� �� ���������"
     Exit Sub
  End If

  TabInd = vbCrLf + String(doc.SelStartCol, vbTab)

  VTName = inputBox("������� ��� ������� ��������:")
  ColumnsList = inputBox("������� ����� ������� ����� �������")
  VTColumns = Split(ColumnsList, ",")

  Txt = VTName + " = �������������(""���������������"");"
  for i = 0 to UBound(VTColumns)
    Txt = Txt + TabInd + VTName + ".������������(""" + Trim(VTColumns(i)) + """);"
  Next

  doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = txt

End sub 'NovayaTabliza()
'
'====================================================
Sub NovayaStroka()

  If Windows.ActiveWnd Is Nothing Then
     MsgBox "��� ��������� ����"
     Exit Sub
  End If
  Set doc = Windows.ActiveWnd.Document
  If doc=docWorkBook Then Set doc=doc.Page(1)
  If doc<>docText Then
     MsgBox "���� �� ���������"
     Exit Sub
  End If

  TabInd = vbCrLf + String(doc.SelStartCol, vbTab)

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
end sub 'NovayaStroka()