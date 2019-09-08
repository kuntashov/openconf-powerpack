@echo off
rem vim:enc=cp866
 
if "-%1"=="-" goto usage
if "%2"=="-b" goto build_installer
	goto generate_nsh

:usage
	echo �ᯮ�짮�����: build.bat [-g \ �����_ᡮન [-b]]
	echo ��ࠬ����:
	echo   �����_ᡮન -- ����� ���ᨨ ᡮન ��ப��, �㤥� �������� � ����� 䠩�� ���⠫���
	echo   -b            -- �᫨ 䫠� �����, � �㤥� �ய�饭� ������� �業�ਥ� ��⠭����
	echo   -g            -- �᫨ 䫠� �����, � �㤥� �ந������� ⮫쪮 ������� �業�ਥ�
	goto exit

:generate_nsh
	echo ������� �業�ਥ� ��⠭���� �ਯ⮢...
	perl tools\gen_nsh.pl scripts\un.intell.nsh.pl ../Intell > scripts/un.Intell.auto.nsh
	perl tools\scripts.nsh.pl -d ..\scripts -f scripts\filters.txt -i scripts > scripts\scripts.auto.nsh
	perl tools\scripts.nsh.pl -u -d ..\scripts -f scripts\filters.txt -i scripts > scripts\un.scripts.auto.nsh	

	echo ������� �業�ਥ� ��⠭���� ��������⮢...
	perl tools\gen_nsh.pl system.nsh.pl ..\system > system.auto.nsh

	rem �ॡ���� ⮫쪮 ������� ����, ��室��
	if "%1"=="-g" goto exit

:build_installer
	echo ���ઠ ���⠫����...
	C:\progra~1\NSIS\makensis.exe /DOC_VerFile=%1 OpenConf.nsi | ..\system\fecho.exe > build.log

:exit
	echo ��⮢�
