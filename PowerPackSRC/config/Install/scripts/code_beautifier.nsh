;===========================================================================
; �������� ��������� ������� code_beautifier.pl

Section "code_beautifier.pl"

	SectionIn 1 2

	!insertmacro OC_IF_PERLSCRIPT_INSTALLED 0 PS_Not_Installed 	
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "${OC_ConfigDir}\scripts\��������������\code_beautifier.pl"
		File "${OC_ConfigDir}\scripts\��������������\RunPerlScripts.vbs"
		goto end

	PS_Not_Installed:
		MessageBox MB_OK "� ����� ������� �� ���������� PerlScript Engine.$\r$\n\
			������ code_beautifier.pl �� ����� ����������!"

	end:

SectionEnd