;===========================================================================
; �������� ��������� ������� navigator.js

Section "navigator.js (���������)"

	SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | navigator.js..."

	SetOutPath "$INSTDIR\config\scripts\���������"
	File "${OC_ConfigDir}\scripts\���������\navigator.js"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\navigatorJS.readme.txt"
	!insertmacro OC_ADD_FILE_TO_STARTMENU "navigatorJS.readme.txt" "navigator.js - �������"

SectionEnd
