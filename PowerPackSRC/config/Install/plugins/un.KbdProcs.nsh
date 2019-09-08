;===========================================================================
; —ценарий деинсталл€ции плагина KbdProcs (MetaEditor)	  

Section "un.KbdProcs"

    SectionIn 1 2

	!insertmacro OC_STATUS "”даление плагинов | KbdProcs..."

    Delete "$INSTDIR\config\KbdProcs.dll"
	Delete "$INSTDIR\config\scripts\KbdProcsHandler.vbs"

SectionEnd