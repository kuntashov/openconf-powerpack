;===========================================================================
; Сценарий установки скрипта intsOLEGenerator.vbs

Section "intsOLEGenerator.vbs"

	SectionIn 1 2

	!insertmacro OC_STATUS "Установка скриптов | intsOLEGenerator.vbs..."

	SetOutPath "$INSTDIR\config\scripts\Intellisence"
	File "${OC_ConfigDir}\scripts\Intellisence\intsOLEGenerator.vbs"

	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\intsOLEGenerator.vbs.txt"
	!insertmacro OC_ADD_FILE_TO_STARTMENU "intsOLEGenerator.vbs.txt" "intsOLEGenerator.vbs - Справка"

SectionEnd