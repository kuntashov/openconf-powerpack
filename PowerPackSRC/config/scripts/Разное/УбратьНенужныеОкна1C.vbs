' � ������ ��������� ���� "������������� ����������", ��� ��������, ���
' "��������� � ���������� �� ������� ��������� � ������",
' ��������� ������� "�������"  � ���� �� ���������� �� ������.
' ������, � ��� ������ :)
'
' ����� ���� ������ �������� �������� ���� - ����� EnumChildWindows
'
' ������: 
'	����� ������� aka artbear	( e-mail:	artbear@bashnet.ru )
'	MetaEditor					( e-mail: shotfire@inbox.ru )
'
Sub Configurator_OnDoModal(Hwnd, Caption, Answer)
	If Caption = "������������� ����������" then
		LB_GETCOUNT		= &H018B
		LB_GETTEXTLEN	= &H018A
		LB_GETTEXT		= &H0189
		
		Set Wrapper = CreateObject("DynamicWrapper")
		Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llsl", "f=s", "r=l"
		ListBoxHandle = Wrapper.FindWindowExA (Hwnd,0,"ListBox",0)
		
		Wrapper.Register "USER32.DLL",   "SendMessage",    "I=llll", "f=s", "r=l"
		ListboxCount = Wrapper.SendMessage (ListBoxHandle, LB_GETCOUNT, 0, 0)
		if ListboxCount <> 2 then ' ��������� ��������, ��� ����� ��� ��������
			exit Sub
		end if
		cnt = Wrapper.SendMessage (ListBoxHandle, LB_GETTEXTLEN, 1, 0)
		if cnt <> Len("��������� � ���������� �� ������� ��������� ������.") then		
			exit Sub ' ����� �� �� ������
		end if
		
		Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
		TextBuff = Space(240)
		
		cnt = Wrapper.SendMessage (ListBoxHandle, LB_GETTEXT, 1, TextBuff)
	    If cnt > 0 Then
	        TextBuff = Left(cstr(TextBuff), cnt)
	    End If

		' ��-�� ���� � DynamicWrapper (�� �� ������!! ) - ��� ��������� ����� 50 �������� �������� ������������ ������ �����
		if TextBuff= "בؑّڑۑܑݑޑߑ�����" then ' ��������� � ���������� �� ������� ��������� � ������
			' ��������� �������� ������ "�������"
			Answer = mbaOK 
		else TextBuff= "��������� � ���������� �� ������� ��������� ������." ' ���� ��� ��� ������ (��� �� ��������)
			' ��������� �������� ������ "�������"
			Answer = mbaOK 
		end if
		
		Wrapper = 0
	end if
End Sub
       
' ���� ���� �� ����
sub GetWindowInfo(hwnd, Group)
	Const WM_GETTEXT = &HD
	
	message Group & "hwnd <" & Hex(hwnd) & ">"
	
	Set Wrapper = CreateObject("DynamicWrapper")
	
 	Wrapper.Register "USER32.DLL",   "SendMessage",    "I=lllr", "f=s", "r=l"
	Title = Space(240)
	cnt = Wrapper.SendMessage(hwnd, WM_GETTEXT ,240,Title) ' (��������� ����)
    If cnt > 0 Then
        Title = Left(cstr(Title), cnt)
		message Group & "Caption <" & Title & ">"
	else
		message Group & "cnt <" & cnt & ">"
    End If
	
 	Wrapper.Register "USER32.DLL",   "GetClassName",    "I=lrl", "f=s", "r=l"
	ClassName = Space(240)
	cnt = Wrapper.GetClassName(hwnd, ClassName, 240) ' ����� ����
    If cnt > 0 Then
        ClassName = Left(cstr(ClassName), cnt)
		message Group & "Class <" & ClassName & ">"
	else
		message Group & "cnt <" & cnt & ">"
    End If
   
	Wrapper = 0
end sub

' ������ �������� �������� ����
sub EnumChildWindows(hwnd, Group)
	
	Set Wrapper = CreateObject("DynamicWrapper")
	                     
	Wrapper.Register "USER32.DLL",   "FindWindowExA",  "I=llll", "f=s", "r=l"
	
	Handle = Wrapper.FindWindowExA (Hwnd,0,0,0)
	while Handle <> 0 
		GetWindowInfo Handle, Group
		EnumChildWindows Handle, Group & "--- "
		Handle = Wrapper.FindWindowExA (hwnd,handle,0,0)
	wend
     
	Wrapper = 0

End Sub