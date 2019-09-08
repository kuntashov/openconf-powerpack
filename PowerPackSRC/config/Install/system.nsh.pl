<<NSH;
;===========================================================================
;; �������� ��������� �����������

Section "���������� (COM-dll, WSC)"
	
	SectionIn 1 2 3
	
	!insertmacro OC_STATUS "��������� ����������� | ����������� ������..."
	
	;; ������� �� �����������
	SetOutPath "\$INSTDIR\\config\\docs"
	File "\${OC_ConfigDir}\\docs\\System.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "System.readme.txt" "���������� - ��������"

	SetOutPath "\$INSTDIR\\config\\system"

	;; ������� fecho.exe ���������� ��������
	File "\${OC_ConfigDir}\\system\\fecho.exe"

	;; COM-dll
@{[`$flst_bat $dir \\.dll\$`]}

	;; ���������� (*.wsc)
@{[`$flst_bat $dir \\.wsc\$`]}

	;; ������� ��� ����������� � ������ ����������� �����������
	File "$dir\\regfiles.js"
	File "$dir\\regall.bat"
	File "$dir\\unregall.bat"

	;; Dynamic Wrapper ��� �� ��������� Win9x
	SetOutPath "\$INSTDIR\\config\\system\\DynWin9x"
	File "\${OC_ConfigDir}\\system\\DynWin9x\\dynwrap.dll"
	File "\${OC_ConfigDir}\\system\\DynWin9x\\readme.txt"

	ClearErrors
	
	!insertmacro OC_STATUS "��������� ����������� | ����������� �����������..."
	
	;; ��������� ������ ����������� �����������
	ExecWait 'wscript.exe //nologo "\$INSTDIR\\config\\system\\regfiles.js" /I /S /L'
	
	;; ���� � ���������� ���������� ������� ��� �������� != 0, �� �������� �� ������
	IfErrors 0 end
		MessageBox MB_OK|MB_ICONEXCLAMATION \\
			"��������� �� ����������� (*.wsc- ��� *.dll-�����)\$\\r\$\\n\\
			�� ���� ���������������� ���������.\$\\r\$\\n\\
			����� ���������� �� ���� ���������������� ����� ������ �� �����\$\\r\$\\n\\
			\$INSTDIR\\config\\system\\regfiles.log.\$\\r\$\\n\\
			��������� ���������� �� ���������� ���� ������ ����� ����� � �����\$\\r\$\\n\\
			\$INSTDIR\\config\\docs\\System.readme.txt"
			
	end:
	
	!insertmacro OC_STATUS "��������� ����������� | ������"
	
SectionEnd ;; "����������"

;===========================================================================
;; �������� ������������� �����������

Section "un.���������� (COM-dll, WSC)"

	!insertmacro OC_STATUS "�������� ����������� | ������ �����������..."

	;; ��������� ������ ������ ����������� �����������
	IfFileExists "\$INSTDIR\\config\\system\\regfiles.js" 0 Delete_Files
		ExecWait 'wscript.exe //nologo "\$INSTDIR\\config\\system\\regfiles.js" /U /S'
		
	Delete_Files:

	!insertmacro OC_STATUS "�������� ����������� | �������� ������..."
	
	;; ������� ��������� �������
	Delete "\$INSTDIR\\config\\system\\regfiles.log"
	Delete "\$INSTDIR\\config\\system\\regfiles.js"
	Delete "\$INSTDIR\\config\\system\\regall.bat"
	Delete "\$INSTDIR\\config\\system\\unregall.bat"

	;; ������� fecho.exe ���������� ��������
	Delete "\$INSTDIR\\config\\system\\fecho.exe"
	
	;; COM-dll
@{[`$flst_bat $dir \\.dll\$ Delete \$INSTDIR\\config\\system`]}

	;; ���������� (*.wsc)
@{[`$flst_bat $dir \\.wsc\$ Delete \$INSTDIR\\config\\system`]}

	Delete "\$INSTDIR\\config\\system\\DynWin9x\\dynwrap.dll"
	Delete "\$INSTDIR\\config\\system\\DynWin9x\\readme.txt"
	RMDir "\$INSTDIR\\config\\system\\DynWin9x"

	;; ������� �������
	Delete "\$INSTDIR\\config\\docs\\System.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "���������� - ��������"

	;; ���� ���������� system ������, ������� ��
	RMDir "\$INSTDIR\\config\\system"	

	!insertmacro OC_STATUS "�������� ����������� | ������"
	
SectionEnd ;; un."����������"
NSH