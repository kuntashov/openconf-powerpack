;===========================================================================
; �������� ������������� ������� author.js

Section "un.author.js (��������� �����������)"

	SectionIn 1

	!insertmacro OC_STATUS "�������� �������� | author.js..."

	Delete "$INSTDIR\config\scripts\��������������\author.js"

	Delete "$INSTDIR\config\docs\authorJS.readme.htm"
	!insertmacro OC_DEL_FILE_FROM_STARTMENU "author.js - �������"

SectionEnd