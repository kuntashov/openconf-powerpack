;===========================================================================
;		СГЕНЕРИРОВАНО АВТОМАТИЧЕСКИ с помощью команды
; 			perl tools\scripts.nsh.pl -u -d ..\scripts -f scripts\filters.txt -i scripts
; 		При следующей генерации ВСЕ ИЗМЕНЕНИЯ БУДУТ ПОТЕРЯНЫ!
; 		Изменения следует вносить в файл tools\scripts.nsh.pl.
;===========================================================================


SubSection un."Ы"
	Section "un.AfterUnInstallЫ"
		RMDir "$INSTDIR\config\scripts\Ы"
	SectionEnd
SubSectionEnd ;; un.Ы
SubSection un."Редактирование"
	!include "scripts\un.author.nsh"
	Section "un.Brackets.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Brackets.vbs ..."
		Delete "$INSTDIR\config\scripts\Редактирование\Brackets.vbs"
	SectionEnd ;; un.Brackets.vbs
	Section "un.mb_utils.js"
		!insertmacro OC_STATUS "Удаление скриптов | mb_utils.js ..."
		Delete "$INSTDIR\config\scripts\Редактирование\mb_utils.js"
	SectionEnd ;; un.mb_utils.js
	Section "un.MultiString.js"
		!insertmacro OC_STATUS "Удаление скриптов | MultiString.js ..."
		Delete "$INSTDIR\config\scripts\Редактирование\MultiString.js"
	SectionEnd ;; un.MultiString.js
	Section "un.refactorer.js"
		!insertmacro OC_STATUS "Удаление скриптов | refactorer.js ..."
		Delete "$INSTDIR\config\scripts\Редактирование\refactorer.js"
	SectionEnd ;; un.refactorer.js
	Section "un.RTrimModule.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | RTrimModule.vbs ..."
		Delete "$INSTDIR\config\scripts\Редактирование\RTrimModule.vbs"
	SectionEnd ;; un.RTrimModule.vbs
	Section "un.Замена кода.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Замена кода.vbs ..."
		Delete "$INSTDIR\config\scripts\Редактирование\Замена кода.vbs"
	SectionEnd ;; un.Замена кода.vbs
	Section "un.Копировать модуль в буфер обмена.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Копировать модуль в буфер обмена.vbs ..."
		Delete "$INSTDIR\config\scripts\Редактирование\Копировать модуль в буфер обмена.vbs"
	SectionEnd ;; un.Копировать модуль в буфер обмена.vbs
	Section "un.Форматирование текста.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Форматирование текста.vbs ..."
		Delete "$INSTDIR\config\scripts\Редактирование\Форматирование текста.vbs"
	SectionEnd ;; un.Форматирование текста.vbs
	Section "un.AfterUnInstallРедактирование"
		RMDir "$INSTDIR\config\scripts\Редактирование"
	SectionEnd
SubSectionEnd ;; un.Редактирование
SubSection un."Разное"
	Section "un.1C++.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | 1C++.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\1C++.vbs"
	SectionEnd ;; un.1C++.vbs
	Section "un.DocPath.js"
		!insertmacro OC_STATUS "Удаление скриптов | DocPath.js ..."
		Delete "$INSTDIR\config\scripts\Разное\DocPath.js"
	SectionEnd ;; un.DocPath.js
	Section "un.ParseCmdLineInConfig.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | ParseCmdLineInConfig.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\ParseCmdLineInConfig.vbs"
	SectionEnd ;; un.ParseCmdLineInConfig.vbs
	Section "un.SetWinCaption.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | SetWinCaption.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\SetWinCaption.vbs"
	SectionEnd ;; un.SetWinCaption.vbs
	Section "un.TelepatSettings.js"
		!insertmacro OC_STATUS "Удаление скриптов | TelepatSettings.js ..."
		Delete "$INSTDIR\config\scripts\Разное\TelepatSettings.js"
	SectionEnd ;; un.TelepatSettings.js
	Section "un.TurboMD_Artur.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | TurboMD_Artur.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\TurboMD_Artur.vbs"
	SectionEnd ;; un.TurboMD_Artur.vbs
	Section "un.Меню всех макросов.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Меню всех макросов.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\Меню всех макросов.vbs"
	SectionEnd ;; un.Меню всех макросов.vbs
	!include "scripts\un.Меню макросов из файла.nsh"
	Section "un.Разработка скриптов.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Разработка скриптов.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\Разработка скриптов.vbs"
	SectionEnd ;; un.Разработка скриптов.vbs
	Section "un.Сохранение открытых окон.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Сохранение открытых окон.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\Сохранение открытых окон.vbs"
	SectionEnd ;; un.Сохранение открытых окон.vbs
	Section "un.УбратьНенужныеОкна1C.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | УбратьНенужныеОкна1C.vbs ..."
		Delete "$INSTDIR\config\scripts\Разное\УбратьНенужныеОкна1C.vbs"
	SectionEnd ;; un.УбратьНенужныеОкна1C.vbs
	Section "un.AfterUnInstallРазное"
		RMDir "$INSTDIR\config\scripts\Разное"
	SectionEnd
SubSectionEnd ;; un.Разное
SubSection un."РаботаСФормами"
	Section "un.Конструкторы Элементов диалога.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы Элементов диалога.vbs ..."
		Delete "$INSTDIR\config\scripts\РаботаСФормами\Конструкторы Элементов диалога.vbs"
	SectionEnd ;; un.Конструкторы Элементов диалога.vbs
	Section "un.Редактор форм.js"
		!insertmacro OC_STATUS "Удаление скриптов | Редактор форм.js ..."
		Delete "$INSTDIR\config\scripts\РаботаСФормами\Редактор форм.js"
	SectionEnd ;; un.Редактор форм.js
	Section "un.Создать кнопку на форме.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Создать кнопку на форме.vbs ..."
		Delete "$INSTDIR\config\scripts\РаботаСФормами\Создать кнопку на форме.vbs"
	SectionEnd ;; un.Создать кнопку на форме.vbs
	Section "un.Создать процедуру и кнопку на форме.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Создать процедуру и кнопку на форме.vbs ..."
		Delete "$INSTDIR\config\scripts\РаботаСФормами\Создать процедуру и кнопку на форме.vbs"
	SectionEnd ;; un.Создать процедуру и кнопку на форме.vbs
	Section "un.Шаблоны форм.js"
		!insertmacro OC_STATUS "Удаление скриптов | Шаблоны форм.js ..."
		Delete "$INSTDIR\config\scripts\РаботаСФормами\Шаблоны форм.js"
	SectionEnd ;; un.Шаблоны форм.js
	Section "un.AfterUnInstallРаботаСФормами"
		RMDir "$INSTDIR\config\scripts\РаботаСФормами"
	SectionEnd
SubSectionEnd ;; un.РаботаСФормами
SubSection un."Общие"
	Section "un.Закрыть окно сообщений.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Закрыть окно сообщений.vbs ..."
		Delete "$INSTDIR\config\scripts\Общие\Закрыть окно сообщений.vbs"
	SectionEnd ;; un.Закрыть окно сообщений.vbs
	Section "un.Клавиатура.js"
		!insertmacro OC_STATUS "Удаление скриптов | Клавиатура.js ..."
		Delete "$INSTDIR\config\scripts\Общие\Клавиатура.js"
	SectionEnd ;; un.Клавиатура.js
	Section "un.Шорткаты.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Шорткаты.vbs ..."
		Delete "$INSTDIR\config\scripts\Общие\Шорткаты.vbs"
	SectionEnd ;; un.Шорткаты.vbs
	Section "un.AfterUnInstallОбщие"
		RMDir "$INSTDIR\config\scripts\Общие"
	SectionEnd
SubSectionEnd ;; un.Общие
SubSection un."Навигация"
	Section "un.FindText.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | FindText.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\FindText.vbs"
	SectionEnd ;; un.FindText.vbs
	Section "un.jumper.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | jumper.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\jumper.vbs"
	SectionEnd ;; un.jumper.vbs
	Section "un.NavigationTools.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | NavigationTools.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\NavigationTools.vbs"
	SectionEnd ;; un.NavigationTools.vbs
	!include "scripts\un.navigator.nsh"
	Section "un.ScriptMethodList.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | ScriptMethodList.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\ScriptMethodList.vbs"
	SectionEnd ;; un.ScriptMethodList.vbs
	Section "un.Навигация.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Навигация.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\Навигация.vbs"
	SectionEnd ;; un.Навигация.vbs
	Section "un.Открыть файл из директивы ЗагрузитьИзФайла.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Открыть файл из директивы ЗагрузитьИзФайла.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\Открыть файл из директивы ЗагрузитьИзФайла.vbs"
	SectionEnd ;; un.Открыть файл из директивы ЗагрузитьИзФайла.vbs
	Section "un.Переходы по модулю.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Переходы по модулю.vbs ..."
		Delete "$INSTDIR\config\scripts\Навигация\Переходы по модулю.vbs"
	SectionEnd ;; un.Переходы по модулю.vbs
	Section "un.AfterUnInstallНавигация"
		RMDir "$INSTDIR\config\scripts\Навигация"
	SectionEnd
SubSectionEnd ;; un.Навигация
SubSection un."Конструкторы"
	Section "un.Конструкторы бухитогов.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы бухитогов.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы бухитогов.vbs"
	SectionEnd ;; un.Конструкторы бухитогов.vbs
	Section "un.Конструкторы документов.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы документов.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы документов.vbs"
	SectionEnd ;; un.Конструкторы документов.vbs
	Section "un.Конструкторы операций.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы операций.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы операций.vbs"
	SectionEnd ;; un.Конструкторы операций.vbs
	Section "un.Конструкторы предопределенных процедур.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы предопределенных процедур.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы предопределенных процедур.vbs"
	SectionEnd ;; un.Конструкторы предопределенных процедур.vbs
	Section "un.Конструкторы справочников.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы справочников.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы справочников.vbs"
	SectionEnd ;; un.Конструкторы справочников.vbs
	Section "un.Конструкторы ТЗ.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Конструкторы ТЗ.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Конструкторы ТЗ.vbs"
	SectionEnd ;; un.Конструкторы ТЗ.vbs
	Section "un.Таблица значений.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Таблица значений.vbs ..."
		Delete "$INSTDIR\config\scripts\Конструкторы\Таблица значений.vbs"
	SectionEnd ;; un.Таблица значений.vbs
	Section "un.AfterUnInstallКонструкторы"
		RMDir "$INSTDIR\config\scripts\Конструкторы"
	SectionEnd
SubSectionEnd ;; un.Конструкторы
SubSection un."ВерсионныйКонтроль"
	Section "un.cvs.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | cvs.vbs ..."
		Delete "$INSTDIR\config\scripts\ВерсионныйКонтроль\cvs.vbs"
	SectionEnd ;; un.cvs.vbs
	Section "un.extforms.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | extforms.vbs ..."
		Delete "$INSTDIR\config\scripts\ВерсионныйКонтроль\extforms.vbs"
	SectionEnd ;; un.extforms.vbs
	Section "un.Сравнить модуль.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Сравнить модуль.vbs ..."
		Delete "$INSTDIR\config\scripts\ВерсионныйКонтроль\Сравнить модуль.vbs"
	SectionEnd ;; un.Сравнить модуль.vbs
	Section "un.СравнитьОбъект.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | СравнитьОбъект.vbs ..."
		Delete "$INSTDIR\config\scripts\ВерсионныйКонтроль\СравнитьОбъект.vbs"
	SectionEnd ;; un.СравнитьОбъект.vbs
	Section "un.AfterUnInstallВерсионныйКонтроль"
		RMDir "$INSTDIR\config\scripts\ВерсионныйКонтроль"
	SectionEnd
SubSectionEnd ;; un.ВерсионныйКонтроль
SubSection un."MD_Tasks"
	Section "un.autoload.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | autoload.vbs ..."
		Delete "$INSTDIR\config\scripts\MD_Tasks\autoload.vbs"
	SectionEnd ;; un.autoload.vbs
	Section "un.ExtractAllReporstIntoExternalFiles.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | ExtractAllReporstIntoExternalFiles.vbs ..."
		Delete "$INSTDIR\config\scripts\MD_Tasks\ExtractAllReporstIntoExternalFiles.vbs"
	SectionEnd ;; un.ExtractAllReporstIntoExternalFiles.vbs
	Section "un.AfterUnInstallMD_Tasks"
		RMDir "$INSTDIR\config\scripts\MD_Tasks"
	SectionEnd
SubSectionEnd ;; un.MD_Tasks
SubSection un."Intellisence"
	Section "un.als2xml.js"
		!insertmacro OC_STATUS "Удаление скриптов | als2xml.js ..."
		Delete "$INSTDIR\config\scripts\Intellisence\als2xml.js"
	SectionEnd ;; un.als2xml.js
	!include "scripts\un.Intellisence.nsh"
	Section "un.Intellisence_DocRef.vbs"
		!insertmacro OC_STATUS "Удаление скриптов | Intellisence_DocRef.vbs ..."
		Delete "$INSTDIR\config\scripts\Intellisence\Intellisence_DocRef.vbs"
	SectionEnd ;; un.Intellisence_DocRef.vbs
	!include "scripts\un.intsOLEGenerator.nsh"
	Section "un.VimComplete.js"
		!insertmacro OC_STATUS "Удаление скриптов | VimComplete.js ..."
		Delete "$INSTDIR\config\scripts\Intellisence\VimComplete.js"
	SectionEnd ;; un.VimComplete.js
	Section "un.AfterUnInstallIntellisence"
		RMDir "$INSTDIR\config\scripts\Intellisence"
	SectionEnd
SubSectionEnd ;; Intellisence
