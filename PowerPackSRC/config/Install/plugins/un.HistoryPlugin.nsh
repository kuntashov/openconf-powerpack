;===========================================================================
; �������� ������������� ������� ������� (����� ������� aka artbear)	  

Section un."�������"
	
	!insertmacro OC_STATUS "�������� �������� | �������..."

    Delete "$INSTDIR\config\historyPlugin.dll"
	Delete "$INSTDIR\config\docs\HistoryPlugin.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "������ ������� - �������"

SectionEnd