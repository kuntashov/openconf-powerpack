' ������� ��� ������ �/��� ���������� ����
'
'	CodeReplace
'	CodeReplaceWithQuestion
'	CodeFramingWithQuestion
'	CodeFramingForTimer
'
' ��������, ���� ����� � ������
'        ���� ��� <> ���������������� �����
' � ��������� ������, ��� ������ ��������� ��
'
' // -- ����� --
' //        ���� ��� <> ���������������� �����
'         ���� ��� <> ���������������� �����
' // -- ����������
'
' ������ "���������� ������ (��� ��������� ���������� �����"
'	CopyLine
'
' ������ "�������� ������� ����� � ������ ����� ���������� � ����������� ��������� ������������ � �������� ����� ����� ����������"
'
'	ExchangeLeftAndRightOfAssign
'
' ��������, ������
' 		�������1 =  ���������������������������[���_�������1];  // ������ 
' �����
' 		���������������������������[���_�������1] =  �������1;  // ������ 
'
' ������� "�������������� �����" � "����������������� �����"
'
'	CommentSelection
'	UnCommentSelection
'
' � ������� �� �������� ������/������� �����������, 
' ���� ���� ������ �� �������� ��� ������� ����� �� ����� ������
' ��������� ������������ vbs-������� (�������� ���������)
'
' ------------------------------------------------------------------------------------------------
' ���� ��� ��� ���������� ������ � ���������.
'
' ������� ������:
' ����� � ����������� ���� 
'      << �����1 = |�����2; >>
' ��� | - ��������� �������,
' 
' �������� �������� ������ ����� �����-������ �������, ��������, "����",
' �������� ������� ������ ��������� ��������� (�����, ������, �����),
' ����� ������, ��������, ������, �������� ��� ������������� � 
'     << �����1 = ������()�����2; >> // ��������� ������, � ��������� ������� ����������� ������
'                               
' ��� �������� �� ����� �������.
' ��� ��� ��������� ��� ��������� ��� ��� �� ����� ��������� �� ������
'     << �����1 = |�����2; >>
' �������� ������
'     << �����1 = ������(�����2);  >> // �������� ������� �����, �� ������ ��
' 
' �.�. ��, ��� � ����������� ������ �� � ����� �������� 
' 
' PS � ��������� ������ ��������� �������� ���, ������� ����� ��������� ������
' �.�. ����� ��������� << �����1 = ������(�����2); >>, ����� Ctrl-Z (��� �������� - ��������),
' ����� ����� �������� ������ << �����1 = ������()�����2); >>
' � ����� �������, ������������ ����� �������� ������ ��� ��������� ��� ������� �������
' ------------------------------------------------------------------------------------------------
'
'
' ������ ����� �������� � ����������� ����������� ��������, 
' ������ �� ����������� ������ ���� �������� ���������
' ���� ��� ���������, ������ �������� � ������� �������
'  
' ������ ��������� � �������� ������������ �������� ����� �����, 
' ��� ��������� �� ���� ����� ����
'
' �����:	����� ������� aka artbear
' e-mail:	artbear@bashnet.ru
'

' ���� true, ����� "�����1 = ������|�����2;" �������� �� "�����1 = ������(�����2)|;"
' ���� false, ����� "�����1 = ������|�����2;" �������� �� "�����1 = ������(|�����2;" - �������, ��� ��� v8
bDontInsertEndCommaInTelepatReplaceFlag = true

sBeginComment = vbCrLf & "// -- ����� -- " ' ������ ����� -  ��������� �������
sEndComment = "// -- ����������" & vbCrLf ' ������ ����� -   ����������� �������

' ������� ��� ������ �/��� ���������� ����
'
' ��������, ���� ����� � ������
'        ���� ��� <> ���������������� �����
' � ��������� ������, ��� ������ ��������� ��
'
' // -- ����� --
' //        ���� ��� <> ���������������� �����
'         ���� ��� <> ���������������� �����
' // -- ����������
'  
' ������ ��������� � �������� ������������ �������� ����� �����, 
' ��� ��������� �� ���� ����� ���� (sBeginComment � sEndComment)
'
Sub CodeReplace()
	CodeReplaceInner sBeginComment, sEndComment, true
End Sub ' CodeReplace 
          
' �� ��, ��� ����������, �� ��� ��������� �������� �������������
Sub CodeReplaceWithQuestion()
	sFrameBegin = InputBox("������� ������ ������������ �����", SelfScript.Name)
	sFrameEnd = InputBox("������� ���������� ������������ �����", SelfScript.Name)
	CodeReplaceInner sFrameBegin, sFrameEnd, true
End Sub ' CodeReplaceWithQuestion

' �������� ������ (��� ���������� ������)
' ��������� �������.
' ��������: ������ ������ 
'	��������� = ���������();
' ����� �������� 
'	���������������("���������");
'	��������� = ���������();
'	����������������("���������");
'
Sub CodeFramingInner(FrameBegin, FrameEnd)
	CodeReplaceInner FrameBegin, FrameEnd, false
End Sub ' CodeFramingInner 

' ������ ��� ������ CodeFramingInner
' �������� ������ (��� ���������� ������)
' ��������� �������.
' ��������: ������ ������ 
'	��������� = ���������();
' ����� �������� 
'	���������������("���������");
'	��������� = ���������();
'	����������������("���������");
'
Sub CodeFramingWithQuestion()
	sFrameBegin = InputBox("������� ������ ������������ �����", SelfScript.Name)
	sFrameEnd = InputBox("������� ���������� ������������ �����", SelfScript.Name)
	CodeFramingInner sFrameBegin, sFrameEnd
End Sub ' CodeFramingWithQuestion

' �������� ������ (��� ���������� ������)
' ������� ��� ��������
' ��������:
'	�������.������("���������");
'	��������� = ���������();
'	�������.����("���������");
Sub CodeFramingForTimer()
	sText = InputBox("������� �������� ������ ��� �������", SelfScript.Name)
	CodeFramingInner "�������.������(""" & sText & """);", "�������.����(""" & sText & """);"
End Sub ' CodeFramingForTimer

' ����� ��� ���������� ���������� ���� (���������� ����� ��� ����� ������, ���� ������ �� ��������)
' ������������ ��������
' ������ ��� ����� ����������������
'
' ����� ����� �������� � ����������� ����������� ��������, 
' ������ �� ����������� ������ ���� �������� ���������
' ���� ��� ���������, ������ �������� � ������� �������
'
Sub CodeReplaceInner(FrameBegin, FrameEnd, bDuplicateBeforeAsComment)
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub

    Line1 = doc.SelStartLine
    Line2 = doc.SelStartLine
    Col1 = 0
    Col2 = doc.LineLen(doc.SelStartLine)

    If (doc.SelStartLine <> doc.SelEndLine) Then ' ���� ��������� �� ���������� �������
        If (doc.SelStartCol = 0) And (doc.SelEndCol = 0) Then ' �������� ����� ���� ��� ��������� �����
            Line2 = doc.SelEndLine - 1
            Col2 = doc.LineLen(doc.SelEndLine - 1)
        Else
            Line2 = doc.SelEndLine
            Col2 = doc.LineLen(doc.SelEndLine)
        End If
    End If
    CurrentText = doc.Range(Line1, Col1, Line2, Col2)

    sArray = Split(CurrentText, vbCrLf)
    CurrentText = FrameBegin & vbCrLf
	if bDuplicateBeforeAsComment then
	    For Each s In sArray
	        CurrentText = CurrentText & "//" & s & vbCrLf
	    Next
	end if
	
    For Each s In sArray
        CurrentText = CurrentText & s & vbCrLf
    Next
    CurrentText = CurrentText & FrameEnd

    doc.Range(Line1, Col1, Line2, doc.LineLen(Line2)) = CurrentText

    ' ����� ������ � �� �� ������� � ������ ����������� �����
	if bDuplicateBeforeAsComment then
	    iNewLine = Line1 + UBound(Split(FrameBegin, vbCrLf)) + UBound(sArray) + 2
	else
	    iNewLine = Line1 + UBound(Split(FrameBegin, vbCrLf)) + 1
	end if
	
    doc.MoveCaret iNewLine, Col1, iNewLine, Col1

End Sub ' CodeReplaceInner 

' ������ "���������� ������ (��� ��������� ���������� �����"
' �������� ������ ��� �����, ��������� ������� �������, 
' ����� ���� �������� ��������� ����� ��� ������ �� ��������, 
' ��������� ����� ���� �������� (�.�. �� ��� ������), 
' ����� ������������ ������.����� � ����� ������ ������ ������ � �� �� �������, � ������� ����� �� ������� �������
'
Sub CopyLine()
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub 
		
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = doc.Range(doc.SelStartLine) & vbCrLf & doc.Range(doc.SelStartLine)
  	else 
		if doc.SelEndLine < doc.LineCount-1 then
    		doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) = doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) & doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0)
		else
    		doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine)) = doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))  & vbCrLf & doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))
		end if
  	end if

    ' ����� ������ � �� �� �������  �� ����� ������ � ������ ����������� �����
    doc.MoveCaret doc.SelEndLine + 1, doc.SelStartCol, doc.SelEndLine + 1, doc.SelStartCol
End Sub ' CopyLine

' ��� ���� ������� (�������)
Sub CopyLine0()
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub
  
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = doc.Range(doc.SelStartLine) & vbCrLf & doc.Range(doc.SelStartLine)
	    ' ����� ������ � �� �� ������� � ������ ����������� �����
	    doc.MoveCaret doc.SelStartLine + 1, doc.SelStartCol, doc.SelStartLine + 1, doc.SelStartCol
		
  	else  ' ���� ��������� �� ���������� �������
	    Line1 = doc.SelStartLine
	    Line2 = doc.SelStartLine
	    Col1 = 0
	    Col2 = doc.LineLen(doc.SelStartLine)
		
        If (doc.SelStartCol = 0) And (doc.SelEndCol = 0) Then ' �������� ����� ���� ��� ��������� �����
            Line2 = doc.SelEndLine - 1
            Col2 = doc.LineLen(doc.SelEndLine - 1)
        Else
            Line2 = doc.SelEndLine
            Col2 = doc.LineLen(doc.SelEndLine)
        End If  
		
	    sArray = Split(doc.Range(Line1, Col1, Line2, Col2), vbCrLf)
		CurrentText = ""
	    For Each s In sArray
	        CurrentText = CurrentText & s & vbCrLf
	    Next
	    For Each s In sArray
	        CurrentText = CurrentText & s & vbCrLf
	    Next
	    doc.Range(Line1, Col1, Line2, doc.LineLen(Line2)) = CurrentText
	
	    ' ����� ������ � �� �� ������� � ������ ����������� �����
	    iNewLine = Line1 + UBound(sArray) + 1
	    doc.MoveCaret iNewLine, doc.SelStartCol, iNewLine, doc.SelStartCol
	
    End If
End Sub ' CopyLine0     

' ��������������� ������� � ExchangeLeftAndRightOfAssign
Function ExchangeLeftAndRightOfAssignInString(Source)
	ExchangeLeftAndRightOfAssignInString = Source

    Set reg = New RegExp
        reg.Pattern = "^(\s*)([^=\s]+)(\s*)=(\s*)([^;]+)(\s*;.*)?"
        reg.IgnoreCase = True
		
    ' ���� �� ������-�����������
    If CommonScripts.RegExpTest("^\s*//", Source) Then
	  	Exit Function
	end if
    Set Matches = reg.Execute(Source)
    If Matches.Count > 0 Then
		ExchangeLeftAndRightOfAssignInString = Matches(0).SubMatches(0) + Matches(0).SubMatches(4) + Matches(0).SubMatches(2) + "=" + Matches(0).SubMatches(3) + Matches(0).SubMatches(1) + Matches(0).SubMatches(5)
	end if
	
End Function

' �������� ������� ����� � ������ ����� ���������� � ����������� ��������� ������������ � �������� ����� ����� ����������
' ��������, ������
' 		�������1 =  ���������������������������[���_�������1];  // ������ 
' �����
' 		���������������������������[���_�������1] =  �������1;  // ������ 

Sub ExchangeLeftAndRightOfAssign()
	
    Set doc = CommonScripts.GetTextDocIfOpened(0)
    If doc Is Nothing Then Exit Sub 
	
	if (doc.SelStartLine = doc.SelEndLine) then
		doc.Range(doc.SelStartLine) = ExchangeLeftAndRightOfAssignInString(doc.Range(doc.SelStartLine))
		
  	else ' �������� ��������� �����
		if doc.SelEndLine < doc.LineCount-1 then
	    	Source = doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0)
		else
	    	Source = doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine))
		end if
	    sArray = Split(Source, vbCrLf)
		Dest = ""
	    For Each s In sArray 
			if s <> "" then
	        	Dest = Dest & ExchangeLeftAndRightOfAssignInString(s) & vbCrLf
			end if
	    Next
		if doc.SelEndLine < doc.LineCount-1 then
			doc.Range(doc.SelStartLine, 0, doc.SelEndLine + 1, 0) = Dest
		else
			doc.Range(doc.SelStartLine, 0, doc.SelEndLine, doc.LineLen(doc.SelEndLine)) = Dest
		end if
  	end if
  	
    doc.MoveCaret doc.SelStartLine, doc.SelStartCol, doc.SelStartLine, doc.SelStartCol
	
End Sub ' ExchangeLeftAndRightOfAssign

' ��� �������� VBScript ������ ����������� ��� ��������
Function GetCommentSymbol(doc)
	GetCommentSymbol = "//"
	If (InStrRev(UCase(doc.Path), ".VBS") > 0) then
		GetCommentSymbol = "'"
	end if
End Function

' �������������� ����� 
' � ������� �� �������� ������ �����������, 
' ���� ���� ������ �� �������� ��� ������� ����� �� ����� ������
' ��������� ������������ vbs-������� (�������� ���������)
'
Sub CommentSelection()
	set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
		                        
	sCommentSymbol = GetCommentSymbol(doc)

	' ���� ������ �� �������� ��� ������� ����� �� ����� ������
	if (doc.SelStartLine = doc.SelEndLine) then
    	Set reg = New RegExp
        reg.Pattern = "^(\s*)(.+)"
		Set Matches = reg.Execute(doc.Range(doc.SelStartLine))
		If Matches.Count > 0 Then
			' ���� ������ �� �����, ����������� ������ ����� ���������� ��������
			doc.Range(doc.SelStartLine) = Matches(0).SubMatches(0) & sCommentSymbol & Matches(0).SubMatches(1)
		else ' ���� �����, ����������� ������ � ������
			doc.Range(doc.SelStartLine) = sCommentSymbol & doc.Range(doc.SelStartLine)
		end if
	    doc.MoveCaret doc.SelStartLine, doc.SelStartCol +Len(sCommentSymbol), doc.SelStartLine, doc.SelEndCol+Len(sCommentSymbol)
	else
  		doc.CommentSel()
  	end if
End Sub
         
' ����������������� ����� 
' � ������� �� �������� ������ �����������, 
' ���� ���� ������ �� �������� ��� ������� ����� �� ����� ������
' ��������� ��������������� vbs-������� (�������� ���������)
'
Sub UnCommentSelection()
  set doc = CommonScripts.GetTextDocIfOpened(0)
  if doc is Nothing then Exit Sub
  	
	sCommentSymbol = GetCommentSymbol(doc)

	' ���� ������ �� �������� ��� ������� ����� �� ����� ������
	if (doc.SelStartLine = doc.SelEndLine) then
		sText = doc.Range(doc.SelStartLine)
		pos = InStr(sText, sCommentSymbol)
		if pos <> 0 then
			doc.Range(doc.SelStartLine) = Left(sText, pos-1) & Mid(sText, pos+Len(sCommentSymbol))
		end if
	    doc.MoveCaret doc.SelStartLine, doc.SelStartCol-Len(sCommentSymbol), doc.SelStartLine, doc.SelEndCol-Len(sCommentSymbol)
	else
		doc.UnCommentSel()
  	end if
End Sub 

Function GetWordFromString(Source)	
	GetWordFromString = ""
	
    Set reg = New RegExp
        reg.Pattern = "^([^\s;]+)([\s;]*)"
        reg.IgnoreCase = True
		
    Set Matches = reg.Execute(Source)
    If Matches.Count > 0 Then
		GetWordFromString = Matches(0).SubMatches(0)
	end if
End Function     
	                  
' ------------------------------------------------------------------------------------------------
' ��� ��� ���������� ������ � ���������.
'
' ������� ������:
' ����� � ����������� ���� 
'      << �����1 = |�����2; >>
' ��� | - ��������� �������,
' 
' �������� �������� ������ ����� �����-������ �������, ��������, "����",
' �������� ������� ������ ��������� ��������� (�����, ������, �����),
' ����� ������, ��������, ������, �������� ��� ������������� � 
'     << �����1 = ������()�����2; >> // ��������� ������, � ��������� ������� ����������� ������
'                               
' ��� �������� �� ����� �������.
' ��� ��� ��������� ��� ��������� ��� ��� �� ����� ��������� �� ������
'     << �����1 = |�����2; >>
' �������� ������
'     << �����1 = ������(�����2);  >> // �������� ������� �����, �� ������ ��
' 
' �.�. ��, ��� � ����������� ������ �� � ����� �������� 
' 
' PS � ��������� ������ ��������� �������� ���, ������� ����� ��������� ������
' �.�. ����� ��������� << �����1 = ������(�����2); >>, ����� Ctrl-Z (��� �������� - ��������),
' ����� ����� �������� ������ << �����1 = ������()�����2); >>
' � ����� �������, ������������ ����� �������� ������ ��� ��������� ��� ������� �������
' ------------------------------------------------------------------------------------------------
'
Sub InsertMethod(InsertType, InsertName, Text)
	
    Set doc = CommonScripts.GetTextDocIfOpened(0)
	if doc is Nothing then Exit Sub
		
	Line = doc.SelStartLine
	Col = doc.SelStartCol
	str1 = doc.Range(Line, 0, Line, Col)
    str2 = doc.Range(Line, Col, Line, doc.LineLen(Line))
	
	str1 = StrReverse(str1)
	 
	str1 = GetWordFromString(str1)
	str1 = StrReverse(str1)
	
	str2 = GetWordFromString(str2)

	if InStr(Text, "!") > 0 then ' ������(!)
		if str2 <> "" then
			if bDontInsertEndCommaInTelepatReplaceFlag then
				Text = Replace(Text, "!", str2)
			else
				replStr = "!)"
				if InStr(1, Text, "!);") > 0 then
					replStr = "!);"
				end if
				Text = Replace(Text, replStr, "!" & str2)				
			end if
		end if
	else ' ��������������()
		if bDontInsertEndCommaInTelepatReplaceFlag then
			Text = Replace(Text, "()", "(" & str2 & ")")
		else
			replStr = "()"
			if InStr(1, Text, "();") > 0 Then
				replStr = "();"
			end if
			Text = Replace(Text, replStr, "(!" & str2)			
		end if
	end if		
	
	doc.Range(Line, Col, Line, Col + Len(str2)) = ""

End Sub

' ���������� ������� ������� ������ �� ������ ����������
' ��������� �������� ����� �������.
' InsertType - ��� ������������ �������� (��������� ����)
' InsertName - ��� ������������ �������� (��� ��� ������� � ������ ����������)
' Text - ����������� �����
' �� ����������� ������ �������������� ����� "!" ���������� ����������
' ������� ����� �������. (�������� ��������� ������ ��� ������������ �������)
' ���� ��������� ������� �� �������, �� �� ��������������� � ����� ������.
' ��� ������� ������� �� ������ ���������� ������ ���������� �� ����������.
' ��� ������� ��������, ��� ������ �,���,�� ��������� �,���,��
'
Sub Telepat_OnInsert(InsertType, InsertName, Text)
	
    Select Case InsertType
        Case 2          ' �������������� ����� �������� ������
			InsertMethod InsertType, InsertName, Text
        Case 3          ' ���������������� ����� �������� ������
			InsertMethod InsertType, InsertName, Text
        Case 7          ' �������������� ����� ����������� ������
			InsertMethod InsertType, InsertName, Text
        Case 8          ' ���������������� ����� ����������� ������
			InsertMethod InsertType, InsertName, Text
        Case 10         ' ����� 1� 
			InsertMethod InsertType, InsertName, Text
    End Select
End Sub

'
' ��������� ������������� �������
'
Sub Init(dummy) ' ��������� ��������, ����� ��������� �� �������� � �������
	
    Set c = Nothing
    On Error Resume Next
    Set c = CreateObject("OpenConf.CommonServices")
    On Error GoTo 0
    If c Is Nothing Then
        Message "�� ���� ������� ������ OpenConf.CommonServices", mRedErr
        Message "������ " & SelfScript.Name & " �� ��������", mInformation
        Scripts.UnLoad SelfScript.Name
		Exit Sub
    End If
    c.SetConfig(Configurator)
	
	' ��� �������� ������� �������������� ���
	c.AddPluginToScript SelfScript, "�������", "Telepat", Telepat
	
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

Init 0
