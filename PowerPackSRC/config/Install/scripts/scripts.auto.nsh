;===========================================================================
;		СГЕНЕРИРОВАНО АВТОМАТИЧЕСКИ с помощью команды
; 			perl tools\scripts.nsh.pl -d ..\scripts -f scripts\filters.txt -i scripts
; 		При следующей генерации ВСЕ ИЗМЕНЕНИЯ БУДУТ ПОТЕРЯНЫ!
; 		Изменения следует вносить в файл tools\scripts.nsh.pl.
;===========================================================================


SubSection "Ы"
SubSectionEnd ;; Ы
SubSection "Редактирование"
	!include "scripts\author.nsh"
	Section "Brackets.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Brackets.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\Brackets.vbs"
	SectionEnd ;; Brackets.vbs
	Section "mb_utils.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | mb_utils.js ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\mb_utils.js"
	SectionEnd ;; mb_utils.js
	Section "MultiString.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | MultiString.js ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\MultiString.js"
	SectionEnd ;; MultiString.js
	Section "refactorer.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | refactorer.js ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\refactorer.js"
	SectionEnd ;; refactorer.js
	Section "RTrimModule.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | RTrimModule.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\RTrimModule.vbs"
	SectionEnd ;; RTrimModule.vbs
	Section "Замена кода.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Замена кода.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\Замена кода.vbs"
	SectionEnd ;; Замена кода.vbs
	Section "Копировать модуль в буфер обмена.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Копировать модуль в буфер обмена.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\Копировать модуль в буфер обмена.vbs"
	SectionEnd ;; Копировать модуль в буфер обмена.vbs
	Section "Форматирование текста.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Форматирование текста.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Редактирование"
		File "..\scripts\Редактирование\Форматирование текста.vbs"
	SectionEnd ;; Форматирование текста.vbs
SubSectionEnd ;; Редактирование
SubSection "Разное"
	Section "1C++.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | 1C++.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\1C++.vbs"
	SectionEnd ;; 1C++.vbs
	Section "DocPath.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | DocPath.js ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\DocPath.js"
	SectionEnd ;; DocPath.js
	Section "ParseCmdLineInConfig.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | ParseCmdLineInConfig.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\ParseCmdLineInConfig.vbs"
	SectionEnd ;; ParseCmdLineInConfig.vbs
	Section "SetWinCaption.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | SetWinCaption.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\SetWinCaption.vbs"
	SectionEnd ;; SetWinCaption.vbs
	Section "TelepatSettings.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | TelepatSettings.js ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\TelepatSettings.js"
	SectionEnd ;; TelepatSettings.js
	Section "TurboMD_Artur.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | TurboMD_Artur.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\TurboMD_Artur.vbs"
	SectionEnd ;; TurboMD_Artur.vbs
	Section "Меню всех макросов.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Меню всех макросов.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\Меню всех макросов.vbs"
	SectionEnd ;; Меню всех макросов.vbs
	!include "scripts\Меню макросов из файла.nsh"
	Section "Разработка скриптов.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Разработка скриптов.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\Разработка скриптов.vbs"
	SectionEnd ;; Разработка скриптов.vbs
	Section "Сохранение открытых окон.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Сохранение открытых окон.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\Сохранение открытых окон.vbs"
	SectionEnd ;; Сохранение открытых окон.vbs
	Section "УбратьНенужныеОкна1C.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | УбратьНенужныеОкна1C.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Разное"
		File "..\scripts\Разное\УбратьНенужныеОкна1C.vbs"
	SectionEnd ;; УбратьНенужныеОкна1C.vbs
SubSectionEnd ;; Разное
SubSection "РаботаСФормами"
	Section "Конструкторы Элементов диалога.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы Элементов диалога.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\РаботаСФормами"
		File "..\scripts\РаботаСФормами\Конструкторы Элементов диалога.vbs"
	SectionEnd ;; Конструкторы Элементов диалога.vbs
	Section "Редактор форм.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Редактор форм.js ..."
		SetOutPath "$INSTDIR\config\scripts\РаботаСФормами"
		File "..\scripts\РаботаСФормами\Редактор форм.js"
	SectionEnd ;; Редактор форм.js
	Section "Создать кнопку на форме.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Создать кнопку на форме.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\РаботаСФормами"
		File "..\scripts\РаботаСФормами\Создать кнопку на форме.vbs"
	SectionEnd ;; Создать кнопку на форме.vbs
	Section "Создать процедуру и кнопку на форме.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Создать процедуру и кнопку на форме.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\РаботаСФормами"
		File "..\scripts\РаботаСФормами\Создать процедуру и кнопку на форме.vbs"
	SectionEnd ;; Создать процедуру и кнопку на форме.vbs
	Section "Шаблоны форм.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Шаблоны форм.js ..."
		SetOutPath "$INSTDIR\config\scripts\РаботаСФормами"
		File "..\scripts\РаботаСФормами\Шаблоны форм.js"
	SectionEnd ;; Шаблоны форм.js
SubSectionEnd ;; РаботаСФормами
SubSection "Общие"
	Section "Закрыть окно сообщений.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Закрыть окно сообщений.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Общие"
		File "..\scripts\Общие\Закрыть окно сообщений.vbs"
	SectionEnd ;; Закрыть окно сообщений.vbs
	Section "Клавиатура.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Клавиатура.js ..."
		SetOutPath "$INSTDIR\config\scripts\Общие"
		File "..\scripts\Общие\Клавиатура.js"
	SectionEnd ;; Клавиатура.js
	Section "Шорткаты.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Шорткаты.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Общие"
		File "..\scripts\Общие\Шорткаты.vbs"
	SectionEnd ;; Шорткаты.vbs
SubSectionEnd ;; Общие
SubSection "Навигация"
	Section "FindText.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | FindText.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\FindText.vbs"
	SectionEnd ;; FindText.vbs
	Section "jumper.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | jumper.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\jumper.vbs"
	SectionEnd ;; jumper.vbs
	Section "NavigationTools.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | NavigationTools.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\NavigationTools.vbs"
	SectionEnd ;; NavigationTools.vbs
	!include "scripts\navigator.nsh"
	Section "ScriptMethodList.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | ScriptMethodList.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\ScriptMethodList.vbs"
	SectionEnd ;; ScriptMethodList.vbs
	Section "Навигация.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Навигация.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\Навигация.vbs"
	SectionEnd ;; Навигация.vbs
	Section "Открыть файл из директивы ЗагрузитьИзФайла.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Открыть файл из директивы ЗагрузитьИзФайла.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\Открыть файл из директивы ЗагрузитьИзФайла.vbs"
	SectionEnd ;; Открыть файл из директивы ЗагрузитьИзФайла.vbs
	Section "Переходы по модулю.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Переходы по модулю.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Навигация"
		File "..\scripts\Навигация\Переходы по модулю.vbs"
	SectionEnd ;; Переходы по модулю.vbs
SubSectionEnd ;; Навигация
SubSection "Конструкторы"
	Section "Конструкторы бухитогов.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы бухитогов.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы бухитогов.vbs"
	SectionEnd ;; Конструкторы бухитогов.vbs
	Section "Конструкторы документов.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы документов.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы документов.vbs"
	SectionEnd ;; Конструкторы документов.vbs
	Section "Конструкторы операций.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы операций.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы операций.vbs"
	SectionEnd ;; Конструкторы операций.vbs
	Section "Конструкторы предопределенных процедур.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы предопределенных процедур.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы предопределенных процедур.vbs"
	SectionEnd ;; Конструкторы предопределенных процедур.vbs
	Section "Конструкторы справочников.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы справочников.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы справочников.vbs"
	SectionEnd ;; Конструкторы справочников.vbs
	Section "Конструкторы ТЗ.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Конструкторы ТЗ.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Конструкторы ТЗ.vbs"
	SectionEnd ;; Конструкторы ТЗ.vbs
	Section "Таблица значений.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Таблица значений.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Конструкторы"
		File "..\scripts\Конструкторы\Таблица значений.vbs"
	SectionEnd ;; Таблица значений.vbs
SubSectionEnd ;; Конструкторы
SubSection "ВерсионныйКонтроль"
	Section "cvs.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | cvs.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\ВерсионныйКонтроль"
		File "..\scripts\ВерсионныйКонтроль\cvs.vbs"
	SectionEnd ;; cvs.vbs
	Section "extforms.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | extforms.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\ВерсионныйКонтроль"
		File "..\scripts\ВерсионныйКонтроль\extforms.vbs"
	SectionEnd ;; extforms.vbs
	Section "Сравнить модуль.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Сравнить модуль.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\ВерсионныйКонтроль"
		File "..\scripts\ВерсионныйКонтроль\Сравнить модуль.vbs"
	SectionEnd ;; Сравнить модуль.vbs
	Section "СравнитьОбъект.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | СравнитьОбъект.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\ВерсионныйКонтроль"
		File "..\scripts\ВерсионныйКонтроль\СравнитьОбъект.vbs"
	SectionEnd ;; СравнитьОбъект.vbs
SubSectionEnd ;; ВерсионныйКонтроль
SubSection "MD_Tasks"
	Section "autoload.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | autoload.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\MD_Tasks"
		File "..\scripts\MD_Tasks\autoload.vbs"
	SectionEnd ;; autoload.vbs
	Section "ExtractAllReporstIntoExternalFiles.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | ExtractAllReporstIntoExternalFiles.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\MD_Tasks"
		File "..\scripts\MD_Tasks\ExtractAllReporstIntoExternalFiles.vbs"
	SectionEnd ;; ExtractAllReporstIntoExternalFiles.vbs
SubSectionEnd ;; MD_Tasks
SubSection "Intellisence"
	Section "als2xml.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | als2xml.js ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\als2xml.js"
	SectionEnd ;; als2xml.js
	!include "scripts\Intellisence.nsh"
	Section "Intellisence_DocRef.vbs"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | Intellisence_DocRef.vbs ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\Intellisence_DocRef.vbs"
	SectionEnd ;; Intellisence_DocRef.vbs
	!include "scripts\intsOLEGenerator.nsh"
	Section "VimComplete.js"
		SectionIn 1 2
		!insertmacro OC_STATUS "Установка скриптов | VimComplete.js ..."
		SetOutPath "$INSTDIR\config\scripts\Intellisence"
		File "..\scripts\Intellisence\VimComplete.js"
	SectionEnd ;; VimComplete.js
SubSectionEnd ;; Intellisence
