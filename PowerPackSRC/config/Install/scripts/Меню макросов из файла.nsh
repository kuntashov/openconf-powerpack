
Section "���� �������� �� �����.vbs"
	
	SectionIn 1 2	

	SetOutPath "$INSTDIR\config\scripts\������"
	File "${OC_ConfigDir}\scripts\������\���� �������� �� �����.vbs"

	SetOutPath "$INSTDIR\config"
	File "${OC_ConfigDir}\Macros.ini"
	File "${OC_ConfigDir}\Macros_all.ini"

SectionEnd