$NAME �������� � ���������� ������
' ������� ��� ��������� � ����������� �������
'
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'
' ����� ������ �������� - "�������� � ���������� ������",
'�.�. ������ ��������� �������� ������� ������ ����� ����������� � �����-�� ������,
'������, ��������, ��� ����������� � �������.
' ���� C����� �������� ������ � ���� ������.
' ����� ������ �����: � ������������ ������������� ��� ����� ��� ���������
'� ������� �������, (��������: ���� ���� ������ ���� ������ ���������),
'����� ����� �������� ������ ����������� �� ��������� ����,
'����� ����������� �������� (�����3), � ��������� ������ ������� �������:
'���� ��� ���������, ���� �������� ������, � ���� ����������.
'������������ ��������� (� ������� �����3) ����������� 2 ������� � ������, ��������, � ������������ � ������������.
'����� ������ ����������� ������� ��������, ���������, �����,
'� ���� ���� ������������, ������ ������ "�������� ������� ������?".
'����� ������ "��", ����� �������� ������ ���������� �� ��������� �����������,��� � ���������.
' PS � �������� �������� ������������ �����3.
' PSS ��������� ������ ��� �������� ��������� � ������ ������� � ����� �������, �������� ������. ������ ������ ���. ������ �� ���� ������� ��������.
' PSSS � ������:
' 1) ����������� �� ������ ������, �� � �����, � ��������.
' 2) ���������/����������� �� ������ 2-� ������, �� � 3 ������ � ���� (���� �����3 - ����� ������ � ����������� ������, ����������)'
'
'
' ��������� �������� ������ ��� ��������:
'%1 - ����, � ������� ����� ������� ����� ��������� ������
'%2 - ����, � ������� ���� ���������
'%3 - ����-��������� ���������/�����������
'
sCmdLineTemplate = "kdiff3 %1 %2 -o %3"

Function GetStringFromTemplate(sTemplate, StringArray)
    sString = sTemplate
    For i = LBound(StringArray) To UBound(StringArray)
        sString = Replace(sString, "%" & CStr(i + 1), """" & StringArray(i) & """")
    Next
    For i = UBound(StringArray) + 1 To 10
        sString = Replace(sString, "%" & CStr(i + 1), "")
    Next
    GetStringFromTemplate = Trim(sString)
End Function ' GetStringFromTemplate

Sub RunDiffWithArray(arr)
    CmdLine = GetStringFromTemplate(sCmdLineTemplate, arr)
'CommonScripts.Echo CmdLine

    CommonScripts.RunCommandAndWait CmdLine, ""
End Sub ' RunDiffWithArray

Function Diff2Files(File1, File2)
    DiffFiles = ""

    arr = Split(File1 & "," & File2, ",")
    RunDiffWithArray arr
End Function ' Diff2Files

Sub TestDiff()
    Diff2Files "W:\����������.txt", "W:\���������� ���.txt"
End Sub

Function Diff3Files(File1, File2, File3)
    DiffFiles3 = ""

    arr = Split(File1 & "," & File2 & "," & File3, ",")
    RunDiffWithArray arr
End Function ' Diff3Files

Function DiffCurrentModuleWithFile(File2)
    DiffCurrentModule = ""

    Set Doc = CommonScripts.GetTextDocIfOpened(0)
    If Doc Is Nothing Then Exit Function

    Set fs = CreateObject("Scripting.FileSystemObject")

'    On Error Resume Next
	Const TemporaryFolder  = 2
	TempFolderName = fs.GetSpecialFolder(TemporaryFolder).Path
	if Right(TempFolderName, 1) <> "\" then
		TempFolderName = TempFolderName + "\"
	end if

    sSourceTempFile = TempFolderName & "tempForDiff_56731.txt"
    sDstTempFile = TempFolderName & "tempForDiff_56732.txt"

    If fs.FileExists(sSourceTempFile) Then fs.DeleteFile sSourceTempFile
    If fs.FileExists(sDstTempFile) Then fs.DeleteFile sDstTempFile

    Doc.SaveToFile sSourceTempFile
'CommonScripts.Echo sSourceTempFile
    Diff3Files sSourceTempFile, File2, sDstTempFile

    If fs.FileExists(sDstTempFile) Then
        If MsgBox("�������� ����� ������ �� ��������� �����������", vbOKCancel + vbQuestion, SelfScript.Name) = vbOK Then
            Doc.LoadFromFile sDstTempFile
        End If
    End If
    If fs.FileExists(sSourceTempFile) Then fs.DeleteFile sSourceTempFile
    If fs.FileExists(sDstTempFile) Then fs.DeleteFile sDstTempFile
End Function ' DiffCurrentModuleWithFile

Sub TestDiffCurrentModule()
    DiffCurrentModuleWithFile "W:\���������� ���.txt"
End Sub

Sub DiffCurrentModuleWithFileSelection()
    Set srv = CreateObject("Svcsvc.Service")
    sDiffFileName = CommonScripts.SelectFileForRead("", "������ 1� � ��������� �����|*.1s;*.txt|������ 1�|*.1s|��������� �����|*.txt|��� �����|*")
    If sDiffFileName = "" Then
        Exit Sub
    End If
    DiffCurrentModuleWithFile sDiffFileName
End Sub 'DiffCurrentModuleWithFileSelection

' ���������� ������ ���� � ����� � ����������� ������ KDiff3
' ���� ������� �������������� ������ � ������� (���� ���������
' ������������� �������� ������, ����� kdiff3 �� �������� � path
Private Function GetKDiff3Folder()
	GetKDiff3Folder = ""
	On Error Resume Next
	GetKDiff3Folder = CStr(CommonScripts.WSH.RegRead("HKCU\Software\KDiff3\"))
	On Error GoTo 0
End Function ' GetKDiff3Folder

Function Text2File(text, path)
    Set textStream = CommonScripts.fso.CreateTextFile(path, 1, 0)
    textStream.Write text
    textStream.Close
    Text2File = path
End Function 'Text2File

Sub DiffSelectionWithClipboard()
    tempPath = CommonScripts.fso.GetSpecialFolder(2).Path
	if Right(tempPath, 1) <> "\" then
		tempPath = tempPath + "\"
	end if
	selFile  = Text2File(CommonScripts.SelectedText, tempPath & "temp1cForDiff_2.txt")
	clipFile = Text2File(CommonScripts.GetFromClipboard, tempPath & "temp1cForDiff_1.txt")

    Diff3Files selFile, clipFile, tempPath & "temp1cForDiff_result.txt"
End Sub 'DiffSelectionWithClipboard

Sub Init(param)
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
	on Error goto 0
	' ��������� ������ ���� � ������������ ����� KDiff3
	tmp = CommonScripts.FSO.BuildPath(GetKDiff3Folder(), "kdiff3.exe")
	If CommonScripts.FSO.FileExists(tmp) Then
		sCmdLineTemplate = Replace(sCmdLineTemplate, "kdiff3", """" & tmp & """")
	End If
End Sub

Init 0
