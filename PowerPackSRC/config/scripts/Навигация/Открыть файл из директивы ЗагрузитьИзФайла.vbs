' ������ ��� �������� - ����� � http://itland.ru/forum/index.php?act=ST&f=37&t=2929&st=
'   ����� - AlexQC
'   ���� � ������������� ����� ��� ������ ���� ��������� #���������������� -
'   ��������� ��������� ����.
'	��� ���� ����, ������������ � "\\" ��� � "�:" (� - ��� �����)- ��������� ����������� �
'	���������� ��� ����.
'   � ��������� ������� ���� ��������� �������� ������������ �������� ����.
'
'         ����������� � ������ ��������� ����� ������� aka artbear
'         ��� e-mail: artbear@bashnet.ru
'         ��� ICQ: 265666057
'
' TODO (a13x)
'	1. �������� ������ �� ������� ���� � ����������� ��������� #����������������
'	2. �������� �������� - ������ ��������� #���������������� ���������� ���������������� �����
'   3. ��� �������� ������� ��� ������� ���������/������ ��� ������� � ������ ��������� #����������������
'		��������� ������������� ��������������� ���� ��� ������


Sub OpenText(name)
    Set fs = CreateObject("Scripting.FileSystemObject")
    If left(name,2)="\\" or mid(name,2,1)=":" then
        ' ���������� ����
        str = name
    Else
        ' ������������� ����
        ' ����������� �� ����� ����� � ����� ���� � ���������, ���� ��� ������� 
        ' ��������� ��� ������� �����, � ����� ����� - ��� ������ ������� ������������
        ' � ����� ���� ���� ��������� ������������ �������� ����
        Set re = New RegExp
        re.Pattern = "\\[\w\s]+\.ert$"
        re.IgnoreCase = true
        Path = re.Replace(DocPath, "")
        If Path = DocPath Then
            ' ��� ������ ����������� ������� ������������, ������� 
            ' ���� ����� �������� ������������ �������� ����
            Path = IBDir
        End If  
        str = CommonScripts.ResolvePath(Path, name)
        'Documents.Open str
    End If
    If Not fs.FileExists(str) Then
        'Set File = fs.CreateTextFile(str, 2)
        'File.Close
        
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
        If d = docText Then ' ������ ������
            d.SaveToFile str
        Else
            If d = docWorkBook Then ' �����
                Set m = d.Page(1) ' ��������� ������
                m.SaveToFile str
            End If
        End If
        
    End If
    Documents.Open str
End Sub ' OpenText

Sub OpenDoc(doc)
    For I = 0 To doc.LineCount-1 ' ���������� ������
        str = trim(doc.range(i,0))  ' & vbCRLF)
        If UCase(left(str,18)) = UCase("#���������������� ") then
            ' ����� �������� - ���������
            str = trim(mid(str,19))
            if left(str,1) = """" then str = mid(str,2)
            if right(str,1) = """" then str = left(str, len(str)-1)
            OpenText str
            Exit Sub
        End If
        If UCase(left(str,14)) = UCase("#LoadFromFile ") then
            ' � �� �� � ����. ���������
            str = trim(mid(str,15))
            if left(str,1) = """" then str = mid(str,2)
            if right(str,1) = """" then str = left(str, len(str)-1)
            OpenText str
            Exit Sub
        End If
    Next
End Sub ' OpenDoc

Sub OpenIncludeFile()
    Set d = CommonScripts.GetTextDoc(0)
    If d Is Nothing Then Exit Sub
    OpenDoc d
End Sub

'TODO
'Sub UnloadDocToExtFile()
'	Set d = CommonScripts.GetTextDoc(0)
'	If d Is Nothing Then Exit Sub
'End Sub

'Sub LoadDocFromExtFile()
'
'End Sub

' ��������� ������������� �������
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
End Sub ' Init

Init 0 ' ��� �������� ������� ��������� �������������
