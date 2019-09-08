;===========================================================================
;		������������� ������������� � ������� �������
; 			perl tools\scripts.nsh.pl -d ..\scripts -f scripts\filters.txt -i scripts
; 		��� ��������� ��������� ��� ��������� ����� ��������!
; 		��������� ������� ������� � ���� tools\scripts.nsh.pl.
;===========================================================================


SubSection "�"
SubSectionEnd ;; �
SubSection "��������������"
	!include "scripts\author.nsh"
	Section "Brackets.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | Brackets.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\Brackets.vbs"
	SectionEnd ;; Brackets.vbs
	Section "mb_utils.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | mb_utils.js ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\mb_utils.js"
	SectionEnd ;; mb_utils.js
	Section "MultiString.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | MultiString.js ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\MultiString.js"
	SectionEnd ;; MultiString.js
	Section "refactorer.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | refactorer.js ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\refactorer.js"
	SectionEnd ;; refactorer.js
	Section "RTrimModule.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | RTrimModule.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\RTrimModule.vbs"
	SectionEnd ;; RTrimModule.vbs
	Section "������ ����.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������ ����.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\������ ����.vbs"
	SectionEnd ;; ������ ����.vbs
	Section "���������� ������ � ����� ������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ���������� ������ � ����� ������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\���������� ������ � ����� ������.vbs"
	SectionEnd ;; ���������� ������ � ����� ������.vbs
	Section "�������������� ������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | �������������� ������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\�������������� ������.vbs"
	SectionEnd ;; �������������� ������.vbs
SubSectionEnd ;; ��������������
SubSection "������"
	Section "1C++.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | 1C++.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\1C++.vbs"
	SectionEnd ;; 1C++.vbs
	Section "DocPath.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | DocPath.js ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\DocPath.js"
	SectionEnd ;; DocPath.js
	Section "ParseCmdLineInConfig.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ParseCmdLineInConfig.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\ParseCmdLineInConfig.vbs"
	SectionEnd ;; ParseCmdLineInConfig.vbs
	Section "SetWinCaption.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | SetWinCaption.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\SetWinCaption.vbs"
	SectionEnd ;; SetWinCaption.vbs
	Section "TelepatSettings.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | TelepatSettings.js ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\TelepatSettings.js"
	SectionEnd ;; TelepatSettings.js
	Section "TurboMD_Artur.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | TurboMD_Artur.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\TurboMD_Artur.vbs"
	SectionEnd ;; TurboMD_Artur.vbs
	Section "���� ���� ��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ���� ���� ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\���� ���� ��������.vbs"
	SectionEnd ;; ���� ���� ��������.vbs
	!include "scripts\���� �������� �� �����.nsh"
	Section "���������� ��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ���������� ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\���������� ��������.vbs"
	SectionEnd ;; ���������� ��������.vbs
	Section "���������� �������� ����.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ���������� �������� ����.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\���������� �������� ����.vbs"
	SectionEnd ;; ���������� �������� ����.vbs
	Section "������������������1C.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������������1C.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������"
		File "..\scripts\������\������������������1C.vbs"
	SectionEnd ;; ������������������1C.vbs
SubSectionEnd ;; ������
SubSection "��������������"
	Section "������������ ��������� �������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ��������� �������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\������������ ��������� �������.vbs"
	SectionEnd ;; ������������ ��������� �������.vbs
	Section "�������� ����.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | �������� ����.js ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\�������� ����.js"
	SectionEnd ;; �������� ����.js
	Section "������� ������ �� �����.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ������ �� �����.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\������� ������ �� �����.vbs"
	SectionEnd ;; ������� ������ �� �����.vbs
	Section "������� ��������� � ������ �� �����.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ��������� � ������ �� �����.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\������� ��������� � ������ �� �����.vbs"
	SectionEnd ;; ������� ��������� � ������ �� �����.vbs
	Section "������� ����.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ����.js ..."
		SetOutPath "$INSTDIR\config\scripts\��������������"
		File "..\scripts\��������������\������� ����.js"
	SectionEnd ;; ������� ����.js
SubSectionEnd ;; ��������������
SubSection "�����"
	Section "������� ���� ���������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ���� ���������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\�����"
		File "..\scripts\�����\������� ���� ���������.vbs"
	SectionEnd ;; ������� ���� ���������.vbs
	Section "����������.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ����������.js ..."
		SetOutPath "$INSTDIR\config\scripts\�����"
		File "..\scripts\�����\����������.js"
	SectionEnd ;; ����������.js
	Section "��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\�����"
		File "..\scripts\�����\��������.vbs"
	SectionEnd ;; ��������.vbs
SubSectionEnd ;; �����
SubSection "���������"
	Section "FindText.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | FindText.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\FindText.vbs"
	SectionEnd ;; FindText.vbs
	Section "jumper.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | jumper.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\jumper.vbs"
	SectionEnd ;; jumper.vbs
	Section "NavigationTools.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | NavigationTools.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\NavigationTools.vbs"
	SectionEnd ;; NavigationTools.vbs
	!include "scripts\navigator.nsh"
	Section "ScriptMethodList.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ScriptMethodList.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\ScriptMethodList.vbs"
	SectionEnd ;; ScriptMethodList.vbs
	Section "���������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ���������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\���������.vbs"
	SectionEnd ;; ���������.vbs
	Section "������� ���� �� ��������� ����������������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ���� �� ��������� ����������������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\������� ���� �� ��������� ����������������.vbs"
	SectionEnd ;; ������� ���� �� ��������� ����������������.vbs
	Section "�������� �� ������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | �������� �� ������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\���������"
		File "..\scripts\���������\�������� �� ������.vbs"
	SectionEnd ;; �������� �� ������.vbs
SubSectionEnd ;; ���������
SubSection "������������"
	Section "������������ ���������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ���������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ���������.vbs"
	SectionEnd ;; ������������ ���������.vbs
	Section "������������ ����������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ����������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ����������.vbs"
	SectionEnd ;; ������������ ����������.vbs
	Section "������������ ��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ��������.vbs"
	SectionEnd ;; ������������ ��������.vbs
	Section "������������ ���������������� ��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ���������������� ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ���������������� ��������.vbs"
	SectionEnd ;; ������������ ���������������� ��������.vbs
	Section "������������ ������������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ������������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ������������.vbs"
	SectionEnd ;; ������������ ������������.vbs
	Section "������������ ��.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������������ ��.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������������ ��.vbs"
	SectionEnd ;; ������������ ��.vbs
	Section "������� ��������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ������� ��������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������"
		File "..\scripts\������������\������� ��������.vbs"
	SectionEnd ;; ������� ��������.vbs
SubSectionEnd ;; ������������
SubSection "������������������"
	Section "cvs.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | cvs.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������������"
		File "..\scripts\������������������\cvs.vbs"
	SectionEnd ;; cvs.vbs
	Section "extforms.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | extforms.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������������"
		File "..\scripts\������������������\extforms.vbs"
	SectionEnd ;; extforms.vbs
	Section "�������� ������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | �������� ������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������������"
		File "..\scripts\������������������\�������� ������.vbs"
	SectionEnd ;; �������� ������.vbs
	Section "��������������.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ��������������.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\������������������"
		File "..\scripts\������������������\��������������.vbs"
	SectionEnd ;; ��������������.vbs
SubSectionEnd ;; ������������������
SubSection "MD_Tasks"
	Section "autoload.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | autoload.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\MD_Tasks"
		File "..\scripts\MD_Tasks\autoload.vbs"
	SectionEnd ;; autoload.vbs
	Section "ExtractAllReporstIntoExternalFiles.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | ExtractAllReporstIntoExternalFiles.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\MD_Tasks"
		File "..\scripts\MD_Tasks\ExtractAllReporstIntoExternalFiles.vbs"
	SectionEnd ;; ExtractAllReporstIntoExternalFiles.vbs
SubSectionEnd ;; MD_Tasks
SubSection "Intellisence"
	Section "als2xml.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | als2xml.js ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\als2xml.js"
	SectionEnd ;; als2xml.js
	!include "scripts\Intellisence.nsh"
	Section "Intellisence_DocRef.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | Intellisence_DocRef.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\Intellisence_DocRef.vbs"
	SectionEnd ;; Intellisence_DocRef.vbs
	!include "scripts\intsOLEGenerator.nsh"
	Section "VimComplete.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "��������� �������� | VimComplete.js ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\VimComplete.js"
	SectionEnd ;; VimComplete.js
SubSectionEnd ;; Intellisence
