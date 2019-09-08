' (c) �.�. �������
' ������ ��������� ���������/������� ����� � �������� ��������
' �� ����� ���������� �������.
' �������� ��� ��-��� WSH, ��� � ��-��� OpenConf
'
'������: $Revision: 1.3 $
'

Set fso = CreateObject("Scripting.FileSystemObject")
Set Shell = CreateObject("WScript.Shell")
InOpenConf = True

On Error Resume Next
Set arg = WScript.Arguments
If Err = 0 Then
    ' ����������� �� ��� wsh
    ' �������� ��� �� ���������� ��������� ������
    InOpenConf = False
    If arg.Count < 3 Then ShowUsage
    Compile = 0
    If LCase(arg(0)) = "/c" Then
        Compile = 1
    ElseIf LCase(arg(0)) = "/d" Then
        Compile = 0
    Else
        ShowUsage
    End If

    Source = arg(1)
    Destination = arg(2)

    Set SourceDir = fso.GetFolder(Source)
    If Err <> 0 Then
        MsgBox "�����-�������� �� ����������", , "ExtForms"
        WScript.Quit
    End If

    Set DestinationDir = fso.GetFolder(Destination)
    If Err <> 0 Then
        Err.Clear
        Set DestinationDir = fso.CreateFolder(Destination)
    End If
    If Err <> 0 Then
        MsgBox "�����-���������� �� ����������", , "ExtForms"
        WScript.Quit
    End If
    if SourceDir.Name = DestinationDir.Name then
        MsgBox "�����-�������� � �����-���������� �� ������ ���������", , "ExtForms"
        WScript.Quit
    End If
    On Error GoTo 0
    If Compile = 0 Then
        DeCompileFolder SourceDir, DestinationDir
    Else
        CompileFolder SourceDir, DestinationDir
    End If
    MsgBox "Done", , "ExtForms"
Else
    ' ����������� ��� ����������. ������ �� ������. ����� ����� �������
    On Error GoTo 0
End If

Sub DeCompileExtForms()
    On Error Resume Next
    sName = InputBox("������� �������-��������", "���������� ��������", IBDir & "ExtForms")
    If Len(sName) = 0 Then ExitSub
    Set SourceDir = fso.GetFolder(sName)
    If Err <> 0 Then
        MsgBox "�������-�������� �� ����������", , "ExtForms"
        Exit Sub
    End If
    dName = InputBox("������� �������-��������", "���������� ��������", sName & "Src")
    If Len(dName) = 0 Then Exit Sub
    Set DestinationDir = fso.GetFolder(dName)
    If Err <> 0 Then
        err.Clear()
        Set DestinationDir = fso.CreateFolder(dName)
    End If
    If Err <> 0 Then
        MsgBox "�������-�������� �� ����������", , "ExtForms"
        Exit Sub
    End If
    On Error GoTo 0
    DeCompileFolder SourceDir, DestinationDir
    MsgBox "Done", , "ExtForms"
    Status ""
End Sub
Sub CompileExtForms()
    On Error Resume Next
    dName = InputBox("������� �������-��������", "���������� ��������", IBDir & "ExtForms")
    If Len(dName) = 0 Then ExitSub
    Set DestinationDir = fso.GetFolder(dName)
    If Err <> 0 Then
        err.Clear()
        Set DestinationDir = fso.CreateFolder(dName)
    End If
    If Err <> 0 Then
        MsgBox "�������-�������� �� ����������", , "ExtForms"
        Exit Sub
    End If
    sName = InputBox("������� �������-��������", "���������� ��������", dName & "Src")
    If Len(sName) = 0 Then Exit Sub
    Set SourceDir = fso.GetFolder(sName)
    If Err <> 0 Then
        MsgBox "�������-�������� �� ����������", , "ExtForms"
        Exit Sub
    End If
    On Error GoTo 0
    CompileFolder SourceDir, DestinationDir
    MsgBox "Done", , "ExtForms"
    Status ""
End Sub

Sub DeCompileFolder(s, d)
    Set sf = s.SubFolders
    For Each sSubFolder In sf
		' ���� ������������ � �������� �������� �����, ����� ��� �������� ���������
		if UCase(d.Path) <> UCase(sSubFolder.Path) Then

	        dName = d.Path & "\" & sSubFolder.Name
	        If Not fso.FolderExists(dName) Then
	            Set dSubFolder = fso.CreateFolder(dName)
	        Else
	            Set dSubFolder = fso.GetFolder(dName)
	        End If
	        DeCompileFolder sSubFolder, dSubFolder
		end if
    Next
    Set sf = s.Files
    For Each file In sf
        If InOpenConf Then Status "���������� " & file.Path & " ..."
        If LCase(Right(file.Name, 4)) = ".ert" Then
            strrun = "gcomp -d -F """ & file.Path & """ -DD """ & d.Path & "\" & file.Name & """"
            Shell.Run strrun, 0, True
        Else
            On Error Resume Next
            fso.CopyFile file.Path, d.Path & "\" & file.Name, True
            If Err <> 0 Then
                If InOpenConf Then
                    Message "�� ������� ����������� " & file.Path, mRedErr
                Else
                    MsgBox "�� ������� ����������� " & file.Path, , "ExtForms"
                End If
            End If
            On Error GoTo 0
        End If
    Next
End Sub

Sub CompileFolder(s, d)
    Set sf = s.SubFolders
    For Each sSubFolder In sf
		If InOpenConf Then Status "������ " & sSubFolder.Path & " ..."
        If LCase(Right(sSubFolder.Name, 4)) = ".ert" Then
            strrun = "gcomp -c -F """ & d.Path & "\" & sSubFolder.Name & """ -DD """ & sSubFolder.Path & """"
            Shell.Run strrun, 0, True
        Else
            dName = d.Path & "\" & sSubFolder.Name
            If Not fso.FolderExists(dName) Then
                Set dSubFolder = fso.CreateFolder(dName)
            Else
                Set dSubFolder = fso.GetFolder(dName)
            End If
            CompileFolder sSubFolder, dSubFolder
        End If
    Next
    Set sf = s.Files
    For Each file In sf
        On Error Resume Next
        fso.CopyFile file.Path, d.Path & "\" & file.Name, True
        If Err <> 0 Then
            If InOpenConf Then
                Message "�� ������� ����������� " & file.Path, mRedErr
            Else
                MsgBox "�� ������� ����������� " & file.Path, , "ExtForms"
            End If
        End If
        On Error GoTo 0
    Next
End Sub

Sub ShowUsage()
    If Not InOpenConf Then
        MsgBox "�������������: extforms.vbs /c|/d SourceFolder DestinationFolder", , "ExtForms"
        WScript.Quit
    Else
        MsgBox "����������� ����������� ������", , "ExtForms"
    End If
End Sub
