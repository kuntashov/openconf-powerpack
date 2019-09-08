' ������ ������������ ��� �������������� ��������
' ���������� ������������ �� ������ ��������� �
' ��������� ������.
' ��� ������������ � ��������� ������ �������
' ������ ���� ������ ���� "/q ����������".
' ��� �������� ���� ������ ���� ��������� � ������.

Dim loadMD  ' ��� ������������ �����
loadMD = ""

Sub Configurator_AllPluginsInit()
    loadMD = InStr(CommandLine, "/q") ' ���� ���� /q
    If loadMD <> 0 Then ' ���� ����, �������� ��� �����
        loadMD = Mid(CommandLine, loadMD + 3)
        '   ���� ����������� ��� ������, ������ ���� ��������� ������
        '   ��� ������ ���� ��������� ��� ��������� �������,
        '   ����� �� �������� �� ����������� �������
        k = 0
        While Scripts.Count <> 1
            n = Scripts.Name(k)
            If LCase(n) <> "autoload" Then
                Message "Unload " & n, mMetaData
                Scripts.Unload k
            Else
                k = 1
            End If
        Wend
        ' �������� ������ ������� ������ ������� ��������� ����.
        ' ���� �� �������.
        While Not Windows.ActiveWnd Is Nothing
        	Windows.ActiveWnd.Close
        Wend
        '   ������ ����� ������� �� �������� ��
        SendCommand cmdLoadMD
    Else
        ' ����� ��������� ������, ��� ��������
        'Message "Not Autoload", mNone
        Scripts.Unload "autoload"
    End If
End Sub

Sub Configurator_OnFileDialog(Saved, Caption, Filter, FileName, Answer)
    '   ����� ������� �� �������� ��, 1� �������� ��� �����.
    If loadMD <> "" Then
        FileName = loadMD
        Answer = mbaOK
    End If
End Sub

Sub Configurator_ConfigWindowCreate()
	MsgBox "AutoLoad Trap Event"
    '   ��� ���� ���������, ������ ��� �����������, ����
    '   ��������� � ��������.
    If loadMD <> "" Then
        Message "Quit", mExclamation3
        Quit True
    End If
End Sub

Sub Configurator_OnMsgBox(Text, Style, DefAnswer, Answer)
    '   �� ����� ����� �������, 1� ������������ ����� ����������
    '   ������ ������ ����, � �������� ������ �� ������ ����������.
    '   �� �� � ���� ��������, ������ �� ��� ������� ����� ��������
    '   �� ���������.
    If loadMD <> "" Then
        Answer = DefAnswer
        Message Text, mExclamation
        '   �������, ���� �� �� ������ �������, �� ������
        '   �������� ����� ���� �������� �� ��, ��� ���������� 1�.
        Text = LCase(Text)
        If InStr(Text, "����") <> 0 And InStr(Text, "�� ������") <> 0 Then Quit False
    End If
End Sub

Sub Configurator_OnDoModal(Hwnd, Caption, Answer)
    '   ���������� ���������� ���������.
    If loadMD <> "" Then
        Message Caption, mInformation
        Answer = mbaOK
    End If
End Sub
