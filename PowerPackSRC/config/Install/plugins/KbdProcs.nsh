;===========================================================================
; Сценарий установки плагина KbdProcs (MetaEditor)	  

Section "KbdProcs"

    SectionIn 1 2

	!insertmacro OC_STATUS "Установка плагинов | KbdProcs..."

    SetOutPath "$INSTDIR\config"
    File "${OC_ConfigDir}\KbdProcs.dll"

	SetOutPath "$INSTDIR\config\scripts"
	File "${OC_ConfigDir}\scripts\KbdProcsHandler.vbs"

SectionEnd