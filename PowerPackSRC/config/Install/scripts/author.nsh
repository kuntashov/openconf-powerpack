;===========================================================================
; �������� ��������� ������� author.js

Section "author.js (��������� �����������)"

	SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | author.js..."

	SetOutPath "$INSTDIR\config\scripts\��������������"
	File "${OC_ConfigDir}\scripts\��������������\author.js"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\authorJS.readme.htm"
	!insertmacro OC_ADD_FILE_TO_STARTMENU "authorJS.readme.htm" "author.js - �������"

SectionEnd