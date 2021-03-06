
	Copyright (c) 2005 OpenConf Community	<http://openconf.itland.ru>
	Copyright (c) 2005 Alexander Kuntashov	<kuntashov@yandex.ru>

	Code Works Helper - редактор файлов-шаблонов для скрипта для массовой
	правки кода Алексея Диркса (CodeIns.pl, CodeWorks.pm).

	Все отзывы/предложения/сообщения об ошибках присылайте на адрес
						kuntashov@yandex.ru
	Просьба добавлять в начало темы письма строку "OC_CWH:" (без кавычек).

СОДЕРЖАНИЕ

	1. Описание
	2. Установка
	3. Использование

1. ОПИСАНИЕ

	Данный файл содержит инструкции по установке и использованию скрипта cwHelper.js
(Code Works Helper), который позволяет создавать шаблоны для скрипта для массовой
правки кода Алексея Диркса (CodeIns.pl, CodeWorks.pm) без необходимости знать язык
программирования perl (пусть даже только базовые моменты).

2. УСТАНОВКА

	1. Поместите папку CodeWorks в каталог config
	2. Зарегистрируйте файл cwProject.wsc (с помощью команды regsvr32)
	3. Скопируйте файл cwHelperLoader.js в каталог config\scripts

3. ИСПОЛЬЗОВАНИЕ

	Редактор запускается с помощью единственного макроса OpenCodeWorksHelper, скрипта
cwHelperLoader.js. После выполнения этого макроса откроется окно редактора файлов шаблона
с пустым проектом.

	Под проектом подразумевается логически связанная последовательность действий, которую
требуется произвести над модулями одного типа в конфигурации, заданная с помощью редактора
шаблонов. Примером таких действий могут служить:

	- создание/удаление объявления глобальной переменной модуля или локальной
		переменной заданной подпрограммы модуля ("Удалить переменную МояПеременная
		из модуля", "Объявить переменную ТекДок в процедуре ПриОткрытии()" и т.п.);

	- создание/удаление объявления подпрограммы (процедуры или функции) с заданным именем;

	- вставка кода в заданные подпрограммы или в исполняемую часть модуля;

	- и так далее.

	По ходу создания шаблона для массовой правки кода, для каждого действия пользователем
задаются параметры, такие как имя объявляемой переменной; имена процедур (создаваемых или
удаляемых, или в которые надо произвести вставку кода); код, который требуется вставить в
модуль/подпрограмму или код создаваемого объявления подпрограммы и т.п.

	Для редактирования проекта (добавления/удаления действий из проекта) служат
кнопки, расположенные в верхней части окна редактора:

	Добавить	-- служит для создания нового действия (элемента проекта). По нажатию
				этой кнопки открывается выпадающее меню со списком доступных для
				добавления в проект действий. При выборе какого-либо действия из этого
				меню открывается диалоговое окно для задания параметров этого действия.

	Удалить		-- служит для удаления отмеченного элемента проекта (или нескольких
				отмеченных элементов проекта сразу)

	Вверх/Вниз	-- не реализовано

	Открыть		-- не реализовано

	Сохранить	-- сохранить шаблон для скрипта для массовой правки кода в файл с заданным
				именем
