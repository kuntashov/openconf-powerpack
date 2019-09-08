;===========================================================================
; �������� ��������� ������� OC2VBA (��������� �������)

Section "OC2VBA (VBA ��� OpenConf)"

    SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | OC2VBA..."

	;; �������� �� ������� �������������� APC
	StrCpy $R0 ""
	ExpandEnvStrings $R0 "%CommonProgramFiles%"

	StrCmp $R0 "" APC_Not_Installed 0
		IfFileExists "$R0\Microsoft Shared\VBA\VBA6\apc6*.dll" APC_Installed APC_Not_Installed

	APC_Not_Installed: ;; APC �� ������ ���� � �������

		MessageBox	MB_OK|MB_ICONINFORMATION \
			"��� ������ ������� OC2VBA ��������� �������$\r$\n\
			�������������� Application Programmability Components.$\r$\n\
			���� ��������� �� ��������� � ��� � �������,$\n$\r\
			������� ������ OC2VBA ���������� �� �����.$\r$\n\
			��������� �� ��������� OC2VBA ����� ��������� � �����$\r$\n\
			$INSTDIR\config\docs\Oc2Vba.readme.txt"

		;; ���� ������� ������������� ������, ��������� ������ ���
		;; �� ��������� ������������, ��� � ��� (�������) �� ������ 
		;; ��������� �� ��������� OC2VBA � ��������� ��������� ���������

		goto Install_Readme

	APC_Installed: ;; APC ����������

	    SetOutPath "$INSTDIR\config"
    	File "${OC_ConfigDir}\oc2vba.dll"

		;; �������� ��� �������������, ��� ��������� "������", �������
		;; ���� ���� ������� ��� ����������, �� ����� ��� ��������
		IfFileExists "$INSTDIR\openconf.ocp" Install_Readme 0
			SetOutPath "$INSTDIR"
			File "${OC_BinDir}\openconf.ocp"

	Install_Readme: 
	
		SetOutPath "$INSTDIR\config\docs"
		File "${OC_ConfigDir}\docs\Oc2Vba.readme.txt"
		!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "Oc2Vba.readme.txt" "Oc2Vba - �������"

SectionEnd ;; OC2VBA
