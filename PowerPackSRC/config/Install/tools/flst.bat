@echo off
rem vim:enc=cp866

if "-%1" == "-" goto usage
if "-%2" == "-" goto need2
if "-%3" == "-" goto need3
if "-%4" == "-" goto need4
	goto run

:need2
	call "%0" %1 .*
	goto exit

:need3
	call "%0" %1 %2 File
	goto exit

:need4
	call "%0" %1 %2 %3 %1
	goto exit

:run
	dir /b /oN %1 | perl -n -e "chomp;/%2/i&&{$p=q(%4)}&&print qq(\t%3 \"$p\\$_\"\n)" 
	goto exit

:usage
	echo �ᯮ�짮�����: flst.bat dir1 [regexp [cmd [dir2]]]
	echo ��ࠬ����:
	echo   dir1   - ��४���, �� ���ன �⡨����� 䠩�� (��᮫��� ��� �⭮�⥫�� ����)
	echo   regexp - ॣ��୮� ��ࠦ����, �ᯮ��㥬�� ��� �⡮� 䠩��� �� ��᪥ (�� 㬮�砭�� .*)
	echo   cmd    - ������� NSIS, ��������� � 䠩���� (File ��� Delete, �� 㬮�砭�� - File)
	echo   dir2   - ���� � 䠩�� �� 楫���� ��設� (����� ��� ��� �����樨 ������� Delete)

:exit

