;===========================================================================
; �������� ������������� ������� ClassWizard (����� ������� aka artbear)	  

Section "un.Classes Wizard"
	
	!insertmacro OC_STATUS "�������� �������� | Classes Wizard..."

    Delete "$INSTDIR\config\ClassesWizard.dll"
    Delete "$INSTDIR\config\scripts\wizard1cpp.vbs"
	Delete "$INSTDIR\config\docs\ClassesWizard.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "Classes Wizard - ��������"

SectionEnd