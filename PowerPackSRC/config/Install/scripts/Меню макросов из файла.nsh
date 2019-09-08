
Section "Меню макросов из файла.vbs"
	
	SectionIn 1 2	

	SetOutPath "$INSTDIR\config\scripts\Разное"
	File "${OC_ConfigDir}\scripts\Разное\Меню макросов из файла.vbs"

	SetOutPath "$INSTDIR\config"
	File "${OC_ConfigDir}\Macros.ini"
	File "${OC_ConfigDir}\Macros_all.ini"

SectionEnd