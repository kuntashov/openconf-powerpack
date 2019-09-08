' ����������/�������������� �������� ���� �������������
' �������� � ������ ����� �����������/�����������������
'
' �����: Orefkov (2004) & artbear (2005)
'
' TODO - ��������, ����� ��������� �������� �� ������������, �.�. ������� ������ ����� ������ ��������, � ����� ��� ����������

' ���������� �������� ���� �������������
Sub SaveOpenDocs()
	
    fName = IBDir & "wnds.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(fName, 2, True, 0)
    f.WriteLine (IsConfigWndOpen())
    Set w = Windows.FirstWnd
    While Not w Is Nothing
        Set d = w.Document
        If d <> docUnknown Then
            If d.ID > 0 And d = docWorkBook Then Set d = d.Page(d.ActivePage)
            stat = CStr(d.ID) & "##" & CStr(d) & "##" & d.Kind & "##" & d.Path & "##"
            If d.ID < 0 And d = docWorkBook Then 
				stat = stat & CStr(d.ActivePage)
			else
				stat = stat & CStr(0)
			end if
			stat = stat & "##"
			
			'bookmark
			Set BookmarksDict = CreateObject("Scripting.Dictionary")
			GetBookmarks d, BookmarksDict
			stat = stat & GetBookmarString(BookmarksDict, "##")

            f.WriteLine (stat)
        End If
        Set w = Windows.NextWnd(w)
    Wend
    f.Close
End Sub

' �������� ����������� ����
Sub OpenSavedDocs()
    fName = IBDir & "wnds.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")

	bUseBookmark = true
	
    Set f = fso.OpenTextFile(fName, 1, False, 0)
    Line = f.ReadLine
    Line = LCase(Line)

    If (Line = "true") Then SendCommand cmdOpenConfigWnd

    While Not f.AtEndOfStream
        Line = f.ReadLine

        del = InStr(Line, "##")
        ID = CLng(Left(Line, del - 1))
        Line = Mid(Line, del + 2)

        del = InStr(Line, "##")
        typedeoc = CLng(Left(Line, del - 1))
        Line = Mid(Line, del + 2)

        del = InStr(Line, "##")
        Kind = Left(Line, del - 1)
        Line = Mid(Line, del + 2)

        del = InStr(Line, "##")
        Path = Left(Line, del - 1)
        Line = Mid(Line, del + 2)

        del = InStr(Line, "##")
		
		if del <=0 then
			bUseBookmark = false
		end if
		
		if bUseBookmark then
	        ActivePage = CLng(Left(Line, del - 1))
	        Line = Mid(Line, del + 2)
	        
	        strBookmark = Line
		end if
		                                                             
		' �� �������������� ��������, ���� �� ��� ������ (����� ������������ ��������)
		if IsNull(CommonScripts.FindOpenDocument(Path)) then
				
	        On Error Resume Next
			iErrNumber = Err.Number
	        Set doc = Nothing
	        If ID > 0 Then
	            Set doc = Documents.DocFromID(ID, typedoc, Path, Kind)
	            If Not doc Is Nothing Then
	                doc.Open 
					iErrNumber = Err.Number
	
					if bUseBookmark then
						BookmarksRestoreInner doc, strBookmark
					end if
	            End If
	        Else
	            Set doc = Documents.Open(Path)
	            If Not doc Is Nothing Then
	                If doc = docWorkBook Then doc.ActivePage = ActivePage 'CLng(Line)
					iErrNumber = Err.Number
	
					if bUseBookmark then
						BookmarksRestoreInner doc, strBookmark
					end if
	            End If
	        End If
	        If iErrNumber <> 0 Then
	            Message "�������� " & Path & " �� ������.", mRedErr
	            Err.Clear
	        End If
			on Error goto 0
		end if
    Wend
End Sub

' �������� ��� �������� � ������� "Scripting.Dictionary"
Sub GetBookmarks(ParamDoc, BookmarksDict)
	If ParamDoc = docWorkBook Then 
		Set doc = ParamDoc.Page(1)
	elseIf ParamDoc = docText Then 
		Set doc = ParamDoc
	else
		Exit Sub
	end if       
	
	iNextLine = 0
	
    On Error Resume Next
		iNextLine = doc.NextBookmark(doc.LineCount-1) ' ��� ��������� �������� � ������ ������
	iErrNumber = Err.Number
    On Error GoTo 0
	if iErrNumber <> 0 then ' ���� ����� ���� ������ � ������
		Exit Sub
	end if
	
	if iNextLine = -1 then ' ����� �� ��������� ������ ����� ��������
		iNextLine = doc.NextBookmark(doc.LineCount-2) ' ��� ��������� �������� � ������ ������
	end if

	while iNextLine <> -1 ' ���� �� ����� ��� �������� 
		if not BookmarksDict.Exists(iNextLine) then 
			BookmarksDict.Add iNextLine, doc.Range(iNextLine)
		
			iNextLine = doc.NextBookmark(iNextLine)
		else
			iNextLine = -1
		end if
	wend
End Sub ' GetBookmarks

Function GetBookmarString(BookmarksDict, Delim)
	Lines = BookmarksDict.Items
	LinesNumber = BookmarksDict.Keys
    s = CStr(BookmarksDict.Count) & Delim
    For i = 0 To BookmarksDict.Count - 1
        s = s & CStr(LinesNumber(i)) & Delim & Lines(i) & Delim
    Next
	GetBookmarString = s
End Function ' GetBookmarString

Sub BookmarksRestoreInner(ParamDoc, s)
	If ParamDoc = docWorkBook Then 
		Set doc = ParamDoc.Page(1)
	elseIf ParamDoc = docText Then 
		Set doc = ParamDoc
	else
		Exit Sub
	end if       
	
	Doc.ClearAllBookMark
	
	Line = s
    del = InStr(Line, "##")
    BookmarksCount = CLng(Left(Line, del - 1))
	Line = Mid(Line, del + 2)
	
	for i = 1 to BookmarksCount
	    del = InStr(Line, "##")
	    BookmarkLineNumber = CLng(Left(Line, del - 1))
		Line = Mid(Line, del + 2)
	    del = InStr(Line, "##")
	    BookmarkText = Left(Line, del - 1)
		Line = Mid(Line, del + 2)
		
		Doc.BookMark(BookmarkLineNumber) = true
	next
End Sub ' BookmarksRestoreInner

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
	SelfScript.AddNamedItem "CommonScripts", c, False
End Sub

Init 0 ' ��� �������� ������� ��������� �������������
