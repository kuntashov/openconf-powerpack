;===========================================================================
; �������� ��������� ������� ������� (����� ������� aka artbear)	  

Section "�������"

    SectionIn 1 2

	!insertmacro OC_STATUS "��������� �������� | �������..."

    SetOutPath "$INSTDIR\config"
    File "${OC_ConfigDir}\historyPlugin.dll"

	;;TODO �������/�������� � ������ ��������� ������ ������� � �����
	;SetOutPath "$INSTDIR\config\docs"
	;File "${OC_ConfigDir}\docs\HistoryPlugin.readme.txt"
	;!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "HistoryPlugin.readme.txt" "������ ������� - �������"

SectionEnd