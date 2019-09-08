;===========================================================================
;; ������� OpenIn1C (������� �����) � ������������� �������

Section "������� OpenIn1C"
	SectionIn 1 2
	!insertmacro OC_STATUS "��������� OpenConf | ������� OpenIn1C..."
	SetOutPath "$INSTDIR\config\system"
	File "${OC_ConfigDir}\system\OpenIn1C.exe"
	File "${OC_ConfigDir}\system\Register_OpenIn1C.vbs"
	SetOutPath "$INSTDIR\config\scripts\������"
	File "${OC_ConfigDir}\scripts\������\�������������������.vbs"
	ExecWait 'wscript.exe	"$INSTDIR\config\system\Register_OpenIn1C.vbs" \
							"$INSTDIR\config\system\OpenIn1C.exe"'
SectionEnd

Section "un.������� OpenIn1C"
	SectionIn 1 2
	!insertmacro OC_STATUS "�������� OpenConf | ������� OpenIn1C..."
	Delete "$INSTDIR\config\system\OpenIn1C.exe"
	Delete "$INSTDIR\config\system\Register_OpenIn1C.vbs"
	Delete "$INSTDIR\config\scripts\������\�������������������.vbs"
	RMDir "$INSTDIR\config\scripts\������"
	RMDir "$INSTDIR\config\scripts"
SectionEnd

;===========================================================================
;; ������������ ���������� �� ���������

Section "������������ ����������"
	SectionIn 1
	SetOutPath "$INSTDIR\config\system"
	File "resources\HotKeys.reg"
	ExecWait 'regedit /s "$INSTDIR\config\system\HotKeys.reg"'
	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\HotKeysList.htm"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "HotKeysList.htm" "������ ������������ ����������"
SectionEnd

Section "un.������������ ����������"
	Delete "$INSTDIR\config\system\HotKeys.reg"	
	Delete "$INSTDIR\config\docs\HotKeysList.htm"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "������ ������������ ����������"
SectionEnd