;===========================================================================
; �������� ������������� ������� Oc2Vba (��������� �������)

Section un."OC2VBA"

	!insertmacro OC_STATUS "�������� �������� | OC2VBA..."

	Delete "$INSTDIR\config\oc2vba.dll"

	!insertmacro OC_USERFRIENDLY_DELETE "$INSTDIR\openconf.ocp"

	Delete  "$INSTDIR\config\docs\Oc2Vba.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "Oc2Vba - �������"

SectionEnd
