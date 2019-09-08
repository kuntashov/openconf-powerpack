;===========================================================================
; Сценарий установки скрипта intsOLEGenerator.vbs

Section "un.intsOLEGenerator.vbs"

	SectionIn 1 2

	!insertmacro OC_STATUS "Удаление скриптов | intsOLEGenerator.vbs..."

	Delete "$INSTDIR\config\scripts\Intellisence\intsOLEGenerator.vbs"

	Delete "$INSTDIR\config\docs\intsOLEGenerator.vbs.txt"
	!insertmacro OC_DEL_FILE_FROM_STARTMENU "intsOLEGenerator.vbs - Справка"

SectionEnd