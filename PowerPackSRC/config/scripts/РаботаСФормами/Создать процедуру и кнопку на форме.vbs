'      ������� ��������� � ������ ��� ��� �� �����
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'
' ��������� ��������� ����� ��������, ����� �������� ������ �������
Dim gsScriptCreateButtonName
gsScriptCreateButtonName = "������� ������ �� �����"

Dim gsProceduresSignature
gsProceduresSignature = "!��������������������������!"
Dim gsFunctionSignature
gsFunctionSignature = "!������������������������!"

' ��������� ��� ������ ������ ��� ������ ���������� ���������
' �� ����� ������� ��� ����������������
Dim DebugFlag '����������� ���������� ����������
'DebugFlag = True '�������� ����� ���������� ���������

' ��������� ������� "������� �������".
' ��������� ��� ������� ������� ����� ���������� ��� 1�.
' ��������� �������� ����� ������, �������������� ��� �� ��������.
' name - ��� ������������ �������.
' text - ����� ������ �������. ����� �������� ���.
' cancel - ���� ������. ��� ��������� � True ������� ������� ����������.
'
Sub Telepat_OnTemplate(Name, Text, Cancel)
		if InStr(Text, gsProceduresSignature) Then
    	Cancel = CreateProceduresAndHisButtonLocal(False, Text)
    Elseif InStr(Text, gsFunctionSignature) then
    	Cancel = CreateProceduresAndHisButtonLocal(True, Text)
    End If
End Sub

Function CreateProceduresAndHisButtonLocal(bIsFunction, sText)
	CreateProceduresAndHisButtonLocal = True      ' �� ��������� ������� ������� �������

'	DebugEcho "sText",  sText
	if bIsFunction then
		sType = "�������"
	else
		sType = "���������"
	end if

	' ����� ������������ �������, ��� ������� ���������/�������
	if not DebugFlag then
		sProcName = InputBox("������� ��� "+sType, "���� "+sType+" � ������ ��� ���")
	else
		sProcName = "���������"
	end if
	If Len(sProcName) = 0 Then Exit Function

	' ��������� � ������ ��������� ��� ����������
	if bIsFunction then
		sText = Replace(sText, gsFunctionSignature, sProcName, 1, -1, 1)
	else
		sText = Replace(sText, gsProceduresSignature, sProcName, 1, -1, 1)
	end if
	DebugEcho "sText",  sText

	set ScriptCreateButton = GetScriptByName(gsScriptCreateButtonName)
	if IsEmpty(ScriptCreateButton) then
		exit Function
	End if

	ScriptCreateButton.CreateButton "�������", sProcName, sProcName+"()", "������_"+sProcName

	CreateProceduresAndHisButtonLocal = False ' �� ����� �������� ������� �������

End Function ' CreateProceduresAndHisButtonLocal

' �������� � ������� ������� ������� ���������� ������
Sub CreateCodeAndHisButtonFromText(bIsFunction, sText)
	set doc = GetTextDoc(null)
	if IsNull(doc) then Exit Sub

	Telepat.ConvertTemplate(sText)

	Cancel = CreateProceduresAndHisButtonLocal(bIsFunction, sText)

	if Cancel then Exit Sub

  ' ���������� ������� �������
  bFindPosition = false
  MyArray = Split(sText, vbCRLF, -1, 1)
  for i=LBound(MyArray) to UBound(MyArray)
  	Pos = inStr(MyArray(i), "<?>")
  	if Pos > 0 then
		  bFindPosition = True
  		exit for
  	end if
  next
  sText = Replace(sText, "<?>", "", 1, -1, 1)

  PosLine = doc.SelStartLine + i
  PosCol = doc.SelStartCol + Pos-1

 	doc.range(doc.SelStartLine,doc.SelStartCol, doc.SelEndLine, doc.SelEndCol) = sText

  if bFindPosition then
  	doc.MoveCaret PosLine, PosCol
  end if

End Sub ' CreateCodeAndHisButtonFromText

' ��� ������ ��� �������������� �������� ��������� � ������
Sub CreateProceduresAndHisButton()
	sText = "//----------------------------- -----------------------"
	sText = sText + vbCrLF + "// "+gsProceduresSignature+"()"
	sText = sText + vbCrLF + "��������� "+gsProceduresSignature+"() �������"
	sText = sText + vbCrLF + vbTab +"<?>"
	sText = sText + vbCrLF + "�������������� // "+gsProceduresSignature+ vbCrLF

	CreateCodeAndHisButtonFromText False, sText
End Sub ' CreateProceduresAndHisButton

' ��� ������ ��� �������������� �������� ��������� � ������
Sub CreateFunctionAndHisButton()
	sText = "//----------------------------- -----------------------"
	sText = sText + vbCrLF + "// "+gsFunctionSignature+"()"
	sText = sText + vbCrLF + "������� "+gsFunctionSignature+"() �������"
	sText = sText + vbCrLF + vbTab +"<?>"
	sText = sText + vbCrLF + "������������ // "+gsFunctionSignature+ vbCrLF

	CreateCodeAndHisButtonFromText True, sText
End Sub ' CreateFunctionAndHisButton

' -----------------------------------------------------------------------------
' �������� ������ �� ��������
' ���� ������ �����������, ������������ empty
'
' ������-�� �����������
'          set GetScript = Scripts(ScriptName)
' �� ����� ��������,  ���� ScriptName - ������� ��������� ����������
'
Function GetScriptByName(ScriptName)
	GetScript = Empty

	bFindScript = False
	for i = 0 to Scripts.Count-1
		if LCase(Scripts.Name(i)) = LCase(ScriptName) then
			bFindScript = True
			exit for
		end if
	next
	if not bFindScript then
		MsgBox("������ ��� ���������� ������ �� ������! " +vbCRLF+vbCRLF+"������ "+""""+sScriptCreateButtonName+".vbs"+""""+" �����������!")
		exit Function
	End if

	DebugEcho "Name",  Scripts.Name(i)

	set GetScriptByName = Scripts.Item(i)
End Function ' GetScriptByName

' ������ ��������, ����� �� ������� � �������
Function GetTextDoc(param)
	GetTextDoc = null

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

	set GetTextDoc = doc
End Function ' GetTextDoc

Sub Echo(text)
	Message text, mNone
End Sub'Echo

Sub Debug(title, msg)
on error resume next
  DebugFlag = DebugFlag
  if err.Number<>0 then
    err.Clear()
    on error goto 0
    Exit Sub
  end if
  if DebugFlag then
    if not (IsEmpty(msg) or IsNull(msg)) then
      msg = CStr(msg)
    end if
    if not (IsEmpty(title) or IsNull(title)) then
      title = CStr(title)
    end if
    If msg="" Then
      Echo(title)
    else
      Echo(title+" - "+msg)
    End If
  End If
on error goto 0
End Sub'Debug

Sub DebugEcho(title, msg)
	Debug title, msg
End Sub'DebugEcho

' ������������� �������. param - ������ ��������,
' ����� �� ������� � �������
'
Sub Init(param)
    Set t = Plugins("�������")  ' �������� ������
    If Not t Is Nothing Then    ' ���� "�������" ��������
        ' ����������� ������ � �������� �������
        SelfScript.AddNamedItem "Telepat", t, False
    Else
        ' ������ �� ��������. �������� � ������
        Scripts.Unload SelfScript.Name
    End If
End Sub

' ��� �������� ������� �������������� ���
Init 0
