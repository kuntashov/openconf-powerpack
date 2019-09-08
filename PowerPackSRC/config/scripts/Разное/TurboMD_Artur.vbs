'������ �������, ������������ ��������� � ����
'������� ������������� ����� � ������
'(���� ������ ������ ���������� ��� ���� �������)
'�� ������� ����.
'��� ���� ����������� ������� ��������� �����/������
'��� ������������� ��������� ������������

Dim BaseDir
BaseDir = IBDir & "unpack\" ' ������� ������� ��� ��������

Dim TurboMdPrmName
TurboMdPrmName = IBDir & "TurboMd.prm"

Dim Collection

' ��������� �������� ����� ���������
Sub MakeDir(Dir)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Left(Dir, 2) = "\\" Then
        'UNC Path
        pos = InStr(3, Dir, "\")    'Server name
        p = Left(Dir, pos)
        Dir = Mid(Dir, pos + 1)
    Else
        p = ""
    End If
    pos = 1
    While pos <> 0
        pos = InStr(Dir, "\")
        If pos = 0 Then
            p = p & Dir
        Else
            p = p & Left(Dir, pos)
            Dir = Mid(Dir, pos + 1)
        End If
        If fso.FolderExists(p) = False Then fso.CreateFolder p
    Wend
End Sub

' ������� ������ �� ������� �������
Sub AddDocToList(ModuleName, FileName)
		if Collection.Exists(ModuleName) then
			Collection.Remove ModuleName
		end if
		Collection.Add ModuleName, FileName
end sub 'AddDocToList

Sub RemoveDocFromList(doc)
		if Collection.Exists(doc.Name) then
			Collection.Remove doc.Name
		end if
End Sub 'RemoveDocFromList()

' ������ ����� turbomd.prm
Sub AnalyzeTurboMDPrm(NullParam) ' �������� �����, ����� �� ��� ����� � ������ ��������
    set Collection = CreateObject("Scripting.Dictionary")

    Set fso = CreateObject("Scripting.FileSystemObject")
    set file = fso.OpenTextFile(TurboMdPrmName, 1, true) '������

		Do While file.AtEndOfStream <> True
      CurrLine = file.ReadLine
      pos = InStr(CurrLine,"=")
      if pos <> 0 then
      	ModuleName = Left(CurrLine, pos-1)
      	FileName = Mid(CurrLine, pos+1)
      	AddDocToList ModuleName, FileName
			end if
		Loop
		file.Close
end sub 'AnalyzeTurboMDPrm

' ���������� ������������ � ���� TurboMD.prm
Sub WriteToTurboMDPrm(NullParam)
    Set fso = CreateObject("Scripting.FileSystemObject")
    set file = fso.CreateTextFile(TurboMdPrmName, 2) '������

		items = Collection.Items
		keys = Collection.Keys
		For i = 0 To Collection.Count -1
			file.WriteLine keys(i) & "=" & items(i)
		Next
		file.Close
end sub 'WriteToTurboMDPrm

'�������� ���������
Sub UnloadDoc(doc)
	'��������� ��� �����
    fName = BaseDir & Replace(doc.Name, ".", "\")
    If doc = docTable Then fName = fName & ".mxl" Else fName = fName & ".txt"
	'�� ����� ����� �������� �������
    lastdec = InStrRev(fName, "\")
    Dir = Left(fName, lastdec - 1)
	'� ������� ���� �������
    MakeDir Dir
	' ��������� �������� � ����
    doc.SaveToFile fName

    AddDocToList doc.Name, "unpack\" & Replace(doc.Name, ".", "\") & ".txt"

End Sub

'���������� ������ ��� �������� ��������� ����
Sub UnloadCurrentWnd()
    Set w = Windows.ActiveWnd
    If w Is Nothing Then
        MsgBox "��� ��������� ����", vbOKOnly, "TurboMD"
        Exit Sub
    End If
    Set d = w.Document
    If d.ID < 2 Then
        MsgBox "���� �� �����, �� ������", vbOKOnly, "TurboMD"
        Exit Sub
    End If

    '������ ����� turbomd.prm
    AnalyzeTurboMDPrm 1

    If d = docText Then     ' ������ ������
        UnloadDoc d
    Else
        If d = docWorkBook Then ' �����
            If MsgBox("��������� �����?", vbQuestion + vbYesNo, "TurboMD") = vbYes Then
            	UnloadDoc d.Page(0) ' ��������� ������
            End If
            If MsgBox("��������� ������?", vbQuestion + vbYesNo, "TurboMD") = vbYes Then
            	UnloadDoc d.Page(1) ' ��������� ������
            End If
        End If
    End If

    WriteToTurboMDPrm 1

End Sub

'���������� ������ ��� �������� ��������� ����
Sub RemoveLinkToCurrentWnd()
    Set w = Windows.ActiveWnd
    If w Is Nothing Then
        MsgBox "��� ��������� ����", vbOKOnly, "TurboMD"
        Exit Sub
    End If
    Set d = w.Document
    If d.ID < 1 Then
        MsgBox "���� �� �����, �� ������", vbOKOnly, "TurboMD"
        Exit Sub
    End If

    '������ ����� turbomd.prm
    AnalyzeTurboMDPrm 1

    If d = docText Then     ' ������ ������
    	RemoveDocFromList d
    Else
        If d = docWorkBook Then ' �����
		    	RemoveDocFromList d.Page(0) ' ��������� ������
    			RemoveDocFromList d.Page(1) ' ��������� ������
        End If
    End If

    WriteToTurboMDPrm 1
End Sub 'RemoveLink_to_CurrentWnd

Sub ClearAllLinks () '����� ��� ����������� �� ��
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(TurboMdPrmName, 2) '������
    file.Write ""
End Sub

'������ ��� �������� ���� ������������� ������ ������� � ������
Sub LoadFromFilesToMD()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(IBDir & "turbomd.prm", 1, True)
    On Error Resume Next
    While f.AtEndOfStream = False
        t = f.ReadLine()
        eq = InStr(t, "=")
        If eq > 0 Then
            dName = Trim(Left(t, eq - 1))
            fName = Trim(Mid(t, eq + 1))
            If Mid(fName, 2, 1) <> ":" And Left(fName, 2) <> "\\" Then fName = IBDir & fName
            Set doc = Documents(dName)
            If Err <> 0 Then
                Message Err.Description, mRedErr
                Err.Clear
            Else
                If doc.LoadFromFile(fName) <> True Then
                    Message "�� ������� ��������� " & doc.Name & " �� " & fName, mBlackErr
                Else
                    Message doc.Name & " �������� �� " & fName, mInformation
                End If
                If Err <> 0 Then
                    Message Err.Description, mRedErr
                    Err.Clear
                End If
            End If
        End If
    Wend
End Sub

' ������ ��� �������� �������� ����� TurboMD.prm
Sub OpenTurboMDPrm()
    Documents.Open IBDir & "turbomd.prm"
End Sub

Sub SaveMD()
    MetaData.SaveMDToFile IBDir & "1cv7new.md", False
End Sub
