$NAME CVS & GComp Script
' (c) �.�. �������
' ����� �������� ��� ������ � gcomp & cvs �� �������������
' ������ ������������ ��������� ��������� �������� ����������:
' ��������� - �� ��� ���������
'    Src	- �������� ������� ����������, ������������� � CVS
'		md	- ����������� GComp'�� ������
'		��������� �����, ������� ����� � ����������, ������ � ������������
'			�������� ��������, ��������, ExtForms, Progs, Classes, Unpuck � ��.
'
'������: $Revision: 1.4 $
'

' ���������
' ��������������, ��� ������� � GComp �������� � PATH. ���� ��� �� ���, ������� ������ ����.
GCompPath = "gcomp.exe "
' ��������������, ��� ������� � CVS �������� � PATH. ���� ��� �� ���, ������� ������ ����.
CVSPath = "cvs.exe "
' ��� �������� ��� ����������� �������
SRCPath = IBDir & "SRC\"
' ����� ����������� ������ CVS � GComp
OutPutTo = 2 ' 0-�� ��������, 1 - � ���� ���������, 2 - ��������� � ����
' �������������� stderr � stdout
ErrToOut = True

' �� �������� ������� ��� �������� ������������ ��� �� ����������
bDontAskAboutConfig = True

' ��� ������ � �����������
Module = ""
'�����������
CVSROOT = ""

' ����� �������
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set svc = CreateObject("Svcsvc.Service")

TempDir = shell.ExpandEnvironmentStrings("%temp%") & "\"
cmdFName = TempDir & "cvsScriptRun.cmd"
outFName = TempDir & "cvsScriptOut.txt"

' ��������� �������� ��� ������ ���������
Set LastOutDoc = Nothing

Sub SplitPath(Path, Dir)
    delim = InStrRev(Path, "\")
    Dir = Mid(Path, delim + 1)
    Path = Left(Path, delim - 1)
End Sub

Function GetRepository(Dummy)
	If Len(CVSROOT) = 0 Then
		GetRepository = InputBox("������� �����������", SelfScript.Name, "")
	Else
		GetRepository = CVSROOT
	End If
End Function

Function GetModule(Dummy)
    If Len(Module) = 0 Then
        If fso.FolderExists(SRCPath & "cvs") Then
            Set rep = fso.OpenTextFile(SRCPath & "cvs\Repository", 1)
            Module = rep.ReadLine()
            rep.Close
        Else
            SplitPath Left(IBDir, Len(IBDir) - 1), Module
            Module = InputBox("������� �������� ������ � �����������", SelfScript.Name, Module)
        End If
    End If
    GetModule = Module
End Function

Function RunCVS(WorkDir, CmdLine)
    ' ��������� ������� cvs � ���������� ����������� ��������� ������
    ' � �������� ��������. ���� ����� ���������������� � ���� cvs.out
    ' �� ��������� ��������. � ����������� �� ����� OutPutTo ���� ����
    ' �� ������������, ���� ��������� � ���� ���������, ���� �����������
    ' � ����.
    ' ��������� ������ ����� ���� �������������.
    ' �� ������ ������ ����������� ������ cvs
    If fso.FileExists(outFName) Then fso.DeleteFile outFName
    ' ��� ������ ���������� ��������� ���� ��� �������
    Set CmdFile = fso.CreateTextFile(cmdFName, True)
    CmdFile.WriteLine "@Echo off"
    WorkDir=svc.AnsiToOEM(WorkDir)
    CmdFile.WriteLine Left(WorkDir, 2)                 ' ������ ����
    CmdFile.WriteLine "cd """ & WorkDir & """"         ' �������� � �������
    if errToOut Then CmdFile.WriteLine "cd | fecho >> """ & outFName & """" ' ������� ������� �������

    CmdLines = Split(CmdLine, vbCrLf)
    For i = 0 To UBound(CmdLines)
        ' ����� cvs ������� ������� ��������, ���������� ������ � OEM
        Line=svc.AnsiToOem(CmdLines(i))
    	if len(Line)>0 Then
	        ' ��� ������ ������ ���������� ������ ������� cvs
	        if errToOut Then
	        	' ������� ��������� ������ �� ����� � � ���� ������
	        	CmdFile.WriteLine "echo " & Line & " | fecho >> """ & outFName & """"
	        End If
	        Line = CVSPath & Line
	        If ErrToOut Then Line = Line & " 2>&1"  ' ������������ stderr � stdout
	        ' ����� ����� ���������� �� ����� � � ����. ��� ��� ������ �������� cvs ������ � ANSI,
	        ' ��� fecho ������� �������� ���������� ����� � OEM
	        Line = Line & " | fecho -o >> """ & outFName & """"
	        CmdFile.WriteLine Line
	    end if
    Next
    CmdFile.Close
    ' ���������� ������ �������
    strrun = "cmd.exe /c """ & cmdFName & """"
    ' �������� �������������� ����
    shell.Run strrun, 1, True ' ��������� �������, ��� �� ���� ���, ����� ���� �������.
    RunCVS = LastOutPut
End Function

Function RunGComp(WorkDir, CmdLine)
    If fso.FileExists(outFName) Then fso.DeleteFile outFName
    ' ��� ������ ���������� ��������� ���� ��� �������
    Set CmdFile = fso.CreateTextFile(cmdFName, True)
    CmdFile.WriteLine "@Echo off"
    WorkDir=svc.AnsiToOEM(WorkDir)
    CmdFile.WriteLine Left(WorkDir, 2)                  ' ������ ����
    CmdFile.WriteLine "cd """ & WorkDir & """"          ' �������� � �������
    if errToOut Then CmdFile.WriteLine "cd | fecho >> """ & outFName & """" ' ������� ������� �������
    CmdLines = Split(CmdLine, vbCrLf)
    For i = 0 To UBound(CmdLines)
        ' ��� ������ ������ ���������� ������ ������� gcomp
        Line=svc.AnsiToOEM(CmdLines(i))
        if len(Line)>0 Then
        	if ErrToOut Then
	        	CmdFile.WriteLine "echo " & Line & " | fecho >> """ & outFName & """"
	        end if
	        Line = GCompPath & Line
	        If ErrToOut Then Line = Line & " 2>&1"  ' ������������ stderr � stdout
	        ' ����� ����� ���������� �� ����� � � ����.
	        Line = Line & " | fecho >> """ & outFName & """"
	        CmdFile.WriteLine Line
	    end if
    Next
    CmdFile.Close
    ' ���������� ������ �������
    strrun = "cmd.exe /c """ & cmdFName & """"
    ' �������� �������������� ����
    shell.Run strrun, 1, True ' ��������� �������, ��� �� ���� ���, ����� ���� �������.
    RunGComp = LastOutPut
End Function


sub RunCVSCmdLine()
    Dir=svc.SelectFolder("������� �������",SRCPath,1,SRCPath)
    if Len(Dir)=0 Then Exit Sub
    cmdLine=InputBox("������� ��������� ������")
    if len(cmdLine)=0 Then Exit Sub
    RunCVS Dir,CmdLine
end sub

Function GetCurrentDocument(Dummy)
    Set GetCurrentDocument = Nothing
    If Windows.ActiveWnd Is Nothing Then
        MsgBox "��� ��������� ����", , SelfScript.Name
        Exit Function
    End If
    Set GetCurrentDocument = Windows.ActiveWnd.Document
End Function

Function DocToSrcPath(doc)
    ' ����� ������ ���� �������, ������� �� ����� ������� ����������
    ' ���������� ���� � ����������� �������
    DocToSrcPath = ""
    If doc = docUnknown Then Exit Function
    If doc.ID < 0 Then Exit Function
    Names = Split(doc.Name, ".")
    Select Case Names(0)
        Case "��������"
            DocToSrcPath = "���������"
        Case "����������"
            DocToSrcPath = "�����������"
    End Select
    If Len(DocToSrcPath) > 0 Then DocToSrcPath = DocToSrcPath + "\" + Names(1)
    'Message DocToSrcPath,mNone
End Function

Sub ShowStatus()
    Set doc = GetCurrentDocument(0)
    If doc Is Nothing Then Exit Sub
    Path = DocToSrcPath(doc)
    If Len(Path) = 0 Then Exit Sub
    If Not fso.FolderExists(SRCPath & Path) Then
        Message "������� " & Path & " �� ����������.", mRedErr
        Exit Sub
    End If
    RunCVS SRCPath, "status """ & Path & """"
End Sub

Sub ImportSrcToRepository()
    If Not fso.FolderExists(SRCPath) Then
        MsgBox "��� �������� � ������������ �������", , SelfScript.Name
        Exit Sub
    End If
    If fso.FolderExists(SRCPath & "cvs") Then
        MsgBox "������� � ������������ ������� ��� ��� ���������", , SelfScript.Name
        Exit Sub
    End If
    msgText = "������������� ������� SRC � �����������?"
    If MsgBox(msgText, vbYesNo, SelfScript.Name) = vbNo Then
        Exit Sub
    End If
    On Error GoTo 0
    ' ��������� ������ ��� �������� ������� SRC
    CVSRepository = GetRepository(0)
    Module = GetModule(0)
    If Len(Module) = 0 Then Exit Sub
    Vendor = InputBox("������� ����� �������������", SelfScript.Name, "vendor")
    Rel = InputBox("������� ����� ������", SelfScript.Name, "release")
    msg = ""
    While Len(msg) = 0
        msg = InputBox("������� �����������", SelfScript.Name, "������ ��� ��������")
        If Len(msg) = 0 Then
            If MsgBox("�������� ������?", vbYesNo, SelfScript.Name) = vbYes Then Exit Sub
        End If
    Wend
    ' ��������� src
    Message "�������� � ������ ������� " & SRCPath, mMetaData
    If Len(CVSRepository) = 0 Then
    	RunCVS SRCPath, "import -C -m """ & msg & """ """ & Module & """ " & Vendor & " " & Rel
    Else
    	RunCVS SRCPath, "-d " & CVSRepository & " import -C -m """ & msg & """ """ & Module & """ " & Vendor & " " & Rel
  	End If
End Sub

Sub DecompileMD()
    fName = IBDir & "1cv7.md"
    If MetaData.Modified > 0 Then
        If MsgBox("���� ���������� �������. ��������� �� ��������� ����?", vbYesNo, SelfScript.Name) = vbNo Then
            Exit Sub
        Else
            fName = IBDir & "1cv7md.tmp"
            Status "���������� ���������� � ���� " & fName & " ..."
            MetaData.SaveMDToFile fName, False
            Status ""
        End If
    End If
    RunGComp IBDir, "-v -d -F """ & fName & """ -D """ & SRCPath & "md"""
End Sub

' �������������� ������������.
Sub InitConfiguration()
	DecompileMD
	DeCompileExtForms
	ImportSrcToRepository
End sub


sub Decomp(s,d,cmdtext)
    Set sf = s.SubFolders
    For Each sSubFolder In sf
        dName = d.Path & "\" & sSubFolder.Name
        If Not fso.FolderExists(dName) Then
            Set dSubFolder = fso.CreateFolder(dName)
        Else
            Set dSubFolder = fso.GetFolder(dName)
        End If
        Decomp sSubFolder, dSubFolder, cmdtext
    Next
    Set sf = s.Files
    For Each file In sf
        If LCase(Right(file.Name, 4)) = ".ert" Then
            strrun = "-d --no-empty-mxl -F """ & file.Path & """ -DD """ & d.Path & "\" & file.Name & """"
            cmdtext=cmdtext & vbCrLf & strrun
        Else
            On Error Resume Next
            fso.CopyFile file.Path, d.Path & "\" & file.Name, True
            If Err <> 0 Then
                Message "�� ������� ����������� " & file.Path, mRedErr
            End If
            On Error GoTo 0
        End If
    Next
end sub

Sub DeCompileExtForms()
	If Not fso.FolderExists(SRCPath) Then MsgBox "������ ���������.",,SelfScript.Name:Exit Sub
	Set Folder=fso.GetFolder(IBDir)
	sel=""
	For Each f In Folder.SubFolders
		if InStr("/new_stru/src/syslog/usrdef/,", "/" & LCase(f.Name) & "/")=0 Then
			if Len(sel)>0 Then sel=sel+vbCrLf
			sel=sel+f.Name
			if fso.FolderExists(SRCPath & f.Name) Then sel=sel & "|c"
		end if
	Next
	if len(sel)=0 Then MsgBox "�� ������� ���������� ���������",,SelfScript.Name:Exit Sub
	sel=svc.SelectValue(sel,"������� ����������� ��������",True)
	if len(sel)=0 Then Exit Sub
	fNames=Split(sel,vbCrLf)
	cmdlines=""
	for i=0 to UBound(fNames)
		Set Source=fso.GetFolder(IBDir & fNames(i))
		dName=SrcPath & fNames(i)
		if fso.FolderExists(dName) Then
			Set Dest=fso.GetFolder(dName)
		else
			set Dest=fso.CreateFolder(dName)
		end if
		Decomp Source,Dest,cmdlines
	next
	if Len(cmdlines)>0 Then RunGComp IBDir, cmdlines
End Sub

Sub CheckOutSrc()
    CheckOutFolder ""
End Sub

Sub CheckOutRefs()
    CheckOutFolder "�����������"
End Sub

Sub CheckOutFolder(Dir)
    Message "��������� ������� ����� � " & SRCPath & Dir & " ...", mMetaData
    If fso.FolderExists(SRCPath & Dir) Then fso.DeleteFolder SRCPath & Dir, True
    Modul = GetModule(0)
    CVSRepository = GetRepository(0)
    If Len(CVSRepository) = 0 Then
    	RunCVS IBDir, "-r co -d """ & SRCPath & Dir & """ """ & Modul & Dir & """"
    Else
    	RunCVS IBDir, "-d" & CVSRepository & " -r co -d """ & SRCPath & Dir & """ """ & Modul & Dir & """"
    End If
End Sub

Sub Configurator_MetaDataSaved(FileName)
	if bDontAskAboutConfig then
    	Exit Sub
    end if

    ' ��������� ��������� GComp'��
    msgText = "��������� ���� GComp'��?"
    If MsgBox(msgText, vbOkCancel, SelfScript.Name) = vbOK Then
        Message "��������� ���� ����������...", mMetaData
        RunGComp IBDir, "-v -d -F """ & FileName & """ -D """ & SRCPath & "md"""
    Else
        Exit Sub
    End If
    ' ��������, ��������� �� ������� SRC ��� ��������� CVS
    If Not fso.FolderExists(SRCPath & "\cvs") Then
        ImportSrcToRepository
    End If
End Sub

Sub BuildMDAndLoad()
	err.Raise 0,SelfScript.Name,"������� �� ��������"
End Sub

Sub UpdateFolder(Dir)
    Message "���������� � �������� " & Dir, mMetaData
    RunCVS Dir, "update"
End Sub

Sub Update()
    Path = svc.SelectFolder("��� �������?", SRCPath,1 + &h4000 , SRCPath)
    UpdateFolder Path
End Sub

Sub Commit()
    Message "�������� ���������", mMetaData
    RunCVS SRCPath, "commit"
End Sub

Sub Configurator_AllPluginsInit()
		If CVSROOT = "" Then
			CVSROOT = shell.ExpandEnvironmentStrings("%$CVSROOT%")
			If CVSROOT = "%$CVSROOT%" Then
				CVSROOT = ""
			End If
		End If
    MdExist=fso.FileExists(IBDir & "1cv7.md")
    SrcExist=fso.FolderExists(SRCPath)
    SrcInCVS=False
    if SrcExist Then SrcInCVS=fso.FolderExists(SRCPath & "cvs")

    If SrcExist Then
    	' ���� ������� �������
        If SrcInCVS Then
        	' �� ��� ���������. ��������� ��������
            Set rep = fso.OpenTextFile(SRCPath & "cvs\Repository", 1)
            Module = rep.ReadLine()
            rep.Close
			if not bDontAskAboutConfig then
	            If MsgBox("�������� SRC �� ����������� " & Module & "?", vbOkCancel, SelfScript.Name) = vbOK Then
	                UpdateFolder SRCPath
	            End If
		    end if
        Else
        	' �� ��� ���������. ��������� ���������
        End If
    Else
    	' ��� �������� �������
        If MdExist Then
        	' ���� ������. ��������� ���������
        Else
        	' ��� �������. ��������� ������� �� �����������
        End If
    End If

' -- ����� -- �������� (��� ����� ��������� ������)
'    Windows.MainWnd.Caption = IBName

End Sub

Sub srvSwitchOutPut()
    OutPutTo = (OutPutTo + 1) Mod 3
    Select Case OutPutTo
        Case 0
            txt = "��� ������"
        Case 1
            txt = "� ���� ���������"
        Case 2
            txt = "� ��������� ����"
    End Select
    Message "������� ����� ������: " & txt, mNone
End Sub

Sub srvSwitchErrToOut()
    ErrToOut = Not ErrToOut
    If ErrToOut Then txt = "" Else txt = " �� "
    Message "��������� �� ������� " & txt & "���������", mNone
End Sub

Function LastOutPut()
    svc.FileO2A outFName
    Set outFile = fso.OpenTextFile(outFName, 1)
    LastOutPut = outFile.ReadAll()
    outFile.Close
    If OutPutTo = 1 Then
        Message LastOutPut, mNone
    ElseIf OutPutTo = 2 Then
        On Error Resume Next
        dType = LastOutDoc
        If Err <> 0 Then Set LastOutDoc = Documents.New(docText)
        On Error GoTo 0
        LastOutDoc.Text = LastOutDoc.Text & vbCrLf & LastOutPut
    End If
End Function

sub DecompileCurrentExtForms()
    If Windows.ActiveWnd Is Nothing Then Exit Sub
    If Windows.ActiveWnd.Document <> docWorkBook Then Exit Sub
    Dim doc ' As WorkBook
    Set doc = Windows.ActiveWnd.Document
    If doc.ID > 0 Then MsgBox "��� �� ������� �����", ,SelfScript.Name: Exit Sub
    If Len(doc.Path) = 0 Then MsgBox "���� �� �������", , SelfScript.Name: Exit Sub
    Path = LCase(doc.Path)
    If InStr(Path, LCase(IBDir)) <> 1 Then
        MsgBox "��������� �������� ������ � �������� � �������� ��", , SelfScript.Name
        Exit Sub
    End If
    doc.Save
    newPath = Replace(Path, IBDir, IBDir & "src\")
    strrun = "-d --no-empty-mxl -F """ & Path & """ -DD """ & newPath & """"
    RunGComp IBDir, strrun
end sub

' �������� ������ ������� ������������ WinCVS
sub RunWinCVS()
	shell.Run """c:\Program Files\GNU\WinCvs 1.3\wincvs.exe""", 1, False
end sub

sub RunWinCVS2()
	shell.Run """c:\Program Files\GNU\WinCvs 2.0\wincvs.exe""", 1, False
end sub
