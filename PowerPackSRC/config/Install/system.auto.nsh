;===========================================================================
;		СГЕНЕРИРОВАНО АВТОМАТИЧЕСКИ с помощью команды
; 			perl tools\gen_nsh.pl system.nsh.pl ..\system
; 		При следующей генерации ВСЕ ИЗМЕНЕНИЯ БУДУТ ПОТЕРЯНЫ!
; 		Изменения следует вносить в шаблон system.nsh.pl.
;===========================================================================

;===========================================================================
;; Сценарий установки компонентов

Section "Компоненты (COM-dll, WSC)"
	
	SectionIn 1 2 3
	
	!insertmacro OC_STATUS "Установка компонентов | Копирование файлов..."
	
	;; Справка по компонентам
	SetOutPath "$INSTDIR\config\docs"
	File "${OC_ConfigDir}\docs\System.readme.txt"
	!insertmacro OC_ADD_DOCFILE_TO_STARTMENU "System.readme.txt" "Компоненты - Описание"

	SetOutPath "$INSTDIR\config\system"

	;; Утилита fecho.exe Александра Орефкова
	File "${OC_ConfigDir}\system\fecho.exe"

	;; COM-dll
	File "..\system\ArtWin.dll"
 	File "..\system\dynwrap.dll"
 	File "..\system\macrosenum.dll"
 	File "..\system\SelectDialog.dll"
 	File "..\system\SelectValue.dll"
 	File "..\system\svcsvc.dll"
 	File "..\system\TLBINF32.DLL"
 	File "..\system\WshExtra.dll"


	;; Скриптлеты (*.wsc)
	File "..\system\1S.StatusIB.wsc"
 	File "..\system\Collections.wsc"
 	File "..\system\CommonServices.wsc"
 	File "..\system\OpenConf.RegistryIniFile.wsc"
 	File "..\system\Registry.wsc"
 	File "..\system\SyntaxAnalysis.wsc"


	;; Скрипты для регистрации и отмены регистрации компонентов
	File "..\system\regfiles.js"
	File "..\system\regall.bat"
	File "..\system\unregall.bat"

	;; Dynamic Wrapper для ОС семейства Win9x
	SetOutPath "$INSTDIR\config\system\DynWin9x"
	File "${OC_ConfigDir}\system\DynWin9x\dynwrap.dll"
	File "${OC_ConfigDir}\system\DynWin9x\readme.txt"

	ClearErrors
	
	!insertmacro OC_STATUS "Установка компонентов | Регистрация компонентов..."
	
	;; Запускаем скрипт регистрации компонентов
	ExecWait 'wscript.exe //nologo "$INSTDIR\config\system\regfiles.js" /I /S /L'
	
	;; Если в результате выполнения команды код возврата != 0, то сообщаем об ошибке
	IfErrors 0 end
		MessageBox MB_OK|MB_ICONEXCLAMATION \
			"Некоторые из компонентов (*.wsc- или *.dll-файлы)$\r$\n\
			не были зарегестрированы корректно.$\r$\n\
			Какие компоненты не были зарегистрированы можно узнать из файла$\r$\n\
			$INSTDIR\config\system\regfiles.log.$\r$\n\
			Подробную информацию об устранении этой ошибки можно найти в файле$\r$\n\
			$INSTDIR\config\docs\System.readme.txt"
			
	end:
	
	!insertmacro OC_STATUS "Установка компонентов | Готово"
	
SectionEnd ;; "Компоненты"

;===========================================================================
;; Сценарий деинсталляции компонентов

Section "un.Компоненты (COM-dll, WSC)"

	!insertmacro OC_STATUS "Удаление компонентов | Отмена регистрации..."

	;; Запускаем скрипт отмены регистрации компонентов
	IfFileExists "$INSTDIR\config\system\regfiles.js" 0 Delete_Files
		ExecWait 'wscript.exe //nologo "$INSTDIR\config\system\regfiles.js" /U /S'
		
	Delete_Files:

	!insertmacro OC_STATUS "Удаление компонентов | Удаление файлов..."
	
	;; Удаляем служебные скрипты
	Delete "$INSTDIR\config\system\regfiles.log"
	Delete "$INSTDIR\config\system\regfiles.js"
	Delete "$INSTDIR\config\system\regall.bat"
	Delete "$INSTDIR\config\system\unregall.bat"

	;; Утилита fecho.exe Александра Орефкова
	Delete "$INSTDIR\config\system\fecho.exe"
	
	;; COM-dll
	Delete "$INSTDIR\config\system\ArtWin.dll"
 	Delete "$INSTDIR\config\system\dynwrap.dll"
 	Delete "$INSTDIR\config\system\macrosenum.dll"
 	Delete "$INSTDIR\config\system\SelectDialog.dll"
 	Delete "$INSTDIR\config\system\SelectValue.dll"
 	Delete "$INSTDIR\config\system\svcsvc.dll"
 	Delete "$INSTDIR\config\system\TLBINF32.DLL"
 	Delete "$INSTDIR\config\system\WshExtra.dll"


	;; Скриптлеты (*.wsc)
	Delete "$INSTDIR\config\system\1S.StatusIB.wsc"
 	Delete "$INSTDIR\config\system\Collections.wsc"
 	Delete "$INSTDIR\config\system\CommonServices.wsc"
 	Delete "$INSTDIR\config\system\OpenConf.RegistryIniFile.wsc"
 	Delete "$INSTDIR\config\system\Registry.wsc"
 	Delete "$INSTDIR\config\system\SyntaxAnalysis.wsc"


	Delete "$INSTDIR\config\system\DynWin9x\dynwrap.dll"
	Delete "$INSTDIR\config\system\DynWin9x\readme.txt"
	RMDir "$INSTDIR\config\system\DynWin9x"

	;; Удаляем справку
	Delete "$INSTDIR\config\docs\System.readme.txt"
	!insertmacro OC_DEL_DOCFILE_FROM_STARTMENU "Компоненты - Описание"

	;; Если директория system пустая, удаляем ее
	RMDir "$INSTDIR\config\system"	

	!insertmacro OC_STATUS "Удаление компонентов | Готово"
	
SectionEnd ;; un."Компоненты"
