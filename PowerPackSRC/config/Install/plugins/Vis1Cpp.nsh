;===========================================================================
; �������� ��������� ������� Visual 1C++ (��������� �������)	  

Section "Visual 1C++ 1.0.0.7"

	SectionIn 1

	!insertmacro OC_STATUS "��������� �������� | Visual 1C++..."

    SetOutPath "$INSTDIR\config"
    File "${OC_ConfigDir}\vis1cpp.dll"

	SetOutPath "$INSTDIR\config\docs"

	File "${OC_ConfigDir}\docs\Vis1Cpp.history.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "Vis1Cpp.history.txt" "Visual 1C++ - ������� ���������"

	File "${OC_ConfigDir}\docs\Vis1Cpp.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "Vis1Cpp.readme.txt" "Visual 1C++ - ��������"

	File "${OC_ConfigDir}\docs\Vis1Cpp.readme.doc"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "Vis1Cpp.readme.doc" "Visual 1C++ - �������"

SectionEnd