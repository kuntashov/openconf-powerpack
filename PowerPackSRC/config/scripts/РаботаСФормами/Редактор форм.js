$NAME Редактор форм
//$Id: Редактор\040форм.js,v 1.4 2006/01/25 12:16:15 alest Exp $
//
//Автор Антипин Алексей aka alest (j_alesant@mail.ru)
//
//распространяется на условиях GPL

//Скрипт позволяет:
//1 добавлять/ удалять / вставлять колонки
	//	функция вставки колонки в выбранную позицию (выбранная колонка №8 станет 9-ой, новая вставится под номером 8)
//2 менять их порядок
	//	при изменении номера колонки в поле номер колонка сдвигает колонки с этим номером и больше вниз
//3 изменять Заголовок, Формулу, Описание, Ширину и Положение
//В комплекте еще должен быть файл bin/config/html/FormEditor.html
function DialogStream (frm) {
	var stream;
	var columns = [];
	this.Columns = function () {
		return columns;
	};
	
	function Column (src, index) {
		//src- строка потока
		if (!src) src = '{"4"," ","70","STATIC","4194","","","","0","2","10","0","0","0","2","","0","270549008","","","","0"}';
		
		var FLAGS = [
					4, // 0- Пропускать при вводе
					16, // 1- Вместо подсказки использовать описание
					32, // 2- Имеет кнопку выбора  
					256, // 3- Растянуть картинку(для объекта Картинка)
					1024, // 4- Картинка по центру
					2048, // 5- Картинка пропорционально
					4096, // 6- Многострочный (для поля ввода)
					8192, // 7- Невидимый
					16384, // 8- Недоступный
					65536, // 9- Прозрачный фон
					131072, // 10- Рамка простая (Для картинки)
					262144, // 11- Рамка выпуклая (Для картинки)
					524288, // 12- Рамка вдавленная (Для картинки)
					1048036, // 13- Автовыбор  (для полей ввода с агрегатным типом)
					2097152, // 14- Неизвестный флаг(появляется у колонок многострочн. части, логику появления отследить не смог)
					4194304, // 15- На следующей строке (положение колонки многострочн. части)
					8388608, // 16- В той же колонке  (положение колонки многострочн. части)
					33554432, // 17- Список с пометками (для списка значений)
					67108864, // 18- Отрицательное красным (для полей ввода типа "Число")
					134217728, // 19- Выводить пиктограммы (Для таблицы значений)
					268435456 // 20- Запретить редактирование.
					];
		var flagValues = new Array(21);

		var string = wrapString(src);
		string = string.replace(/("),(")/g, "$1\r\n$2");
		var columnParams = string.split("\r\n");
		
		this.Src = function (newSrc) {
			if (newSrc) src = newSrc;
			return src;
		};
		
		function splitFlags () {
			var n = unquote(columnParams[17]);
			for (var i = FLAGS.length; i > -1; i--) {
				if (n >= FLAGS[i]) {
					n -= FLAGS[i];
					flagValues[i] = 1;
				}
				else {
					flagValues[i] = 0;
				}
			}
		};
		
		function joinFlags () {
			var n = 0;
			for (var i = 0; i < flagValues.length; i++) {
				if (flagValues[i]) {
					n += FLAGS[i];
					//Message(i + "="+FLAGS[i]);
				}
			}
			return n;
		};
		
		splitFlags();

		function getFlag (id) {
			var value;
			if (id == "location") {
				if (flagValues[15]) {
					value = 2;
				}
				else if (flagValues[16]) {
					value = 3;
				}
				else
					value = 1;
			}
			else if (id == "invisible") {
				if (flagValues[7])
					value = 1;
				else
					value = 0;
			}
			return value;
		};
		
		function updateFlag (id, value) {
			if (id == "location") {
				flagValues[15] = 0;
				flagValues[16] = 0;
				if (value == 2) {
					flagValues[15] = 1;
				}
				else if (value == 3) {
					flagValues[16] = 1;
				}
			}
			else if (id == "invisible") {
				flagValues[7] = value;
			}
		};
		
		function unquote(str) {
			return str.substr(1, str.length-2);
		};

		function quote(str) {
			return "\""+str+"\"";
		};
			
		function wrapAttr (attr) {
			attr = attr.replace(/""{([^{(),"}]+)}""/g, "sendkey|$1|"); //EmulationKey(""{RIGHT}"",0);
			attr = attr.replace(/\"/g, "k2av");
			attr = attr.replace(/\r\n/g, "s5ace");

			return attr;
		};
		
		function unwrapAttr (attr) {
			attr = attr.replace(/k2av/g, "\"");
			attr = attr.replace(/s5ace/g, "\r\n");
			attr = attr.replace(/sendkey\|([^|]+)\|/g, "\"{$1}\"");

			return attr;
		};
					
		this.attrs = {
					'index': index,
					id: unquote(columnParams[7]),
					caption:unquote(columnParams[1]),
					formula: unquote(unwrapAttr(columnParams[6])),
					help: unquote(unwrapAttr(columnParams[19])),
					tip: unquote(unwrapAttr(columnParams[20])),
					width: unquote(columnParams[2]),
					location: getFlag("location"),
					invisible: getFlag("invisible")
					};

		this.update = function (attrs) {
			for (var attrId in attrs) {
				this.attrs[attrId] = attrs[attrId];
			}
			columnParams[7] = quote(this.attrs['id']);
			columnParams[1] = quote(this.attrs['caption']);
			columnParams[6] = quote(wrapAttr(this.attrs['formula']));
			columnParams[19] = quote(wrapAttr(this.attrs['help']));
			columnParams[20] = quote(wrapAttr(this.attrs['tip']));
			columnParams[2] = quote(this.attrs['width']);
			
			updateFlag("location", this.attrs["location"]);
			updateFlag("invisible", this.attrs["invisible"]);
			columnParams[17] = quote(joinFlags());
			this.Src(columnParams.join(","));
		};

		this.isLabel = function () {
			if (columnParams[0] == "{\"4\"")
				return 1;
			else
				return 0;
		};
		
	};

	function wrapString (string) {	//спасибо Сергею Гурову
		string = string.replace(/""{([^{(),"}]+)}""/g, "sendkey|$1|"); //EmulationKey(""{RIGHT}"",0);
		string = string.replace(/(,|\()"""",/g, "$1\k2avk2av,"); // fix bug with Раскрой.frm and others with [,""""",] inside atom #1880
								// and глЗаполнитьНумерацию("""", to #3012
		
		string = string.replace(/""([^{,"}])/g, "k2av$1");
		string = string.replace(/([^{,"}])""/g, "$1k2av");
		string = string.replace(/ k2av"",/g, "k2avk2av");
		string = string.replace(/^(\s*?)\r\n/gm, "$1s5ace");
		string = string.replace(/\r\n([^{])/gm, "s5ace$1");
	
		return string;
	};
	
	function unwrapString (string) {
		string = string.replace(/k2av/g, "\"\"");
		string = string.replace(/s5ace/g, "\r\n");
		string= string.replace(/sendkey\|([^|]+)\|/g, "\"{$1}\"");

		return string;
	};
	
	function refreshIndices () {
		for (var i = 0; i < columns.length; i++) {
			columns[i].attrs['index'] = i;
		}
	};
	
	function checkColumnId (index) {
		var id = columns[index].attrs.id;
		if (id == "") return;
		
		var string = "=";
		for (var i = 0; i < columns.length; i++) {
			if (i != index)
				string += columns[i].attrs.id + "=";
		}
		while (string.indexOf("=" + id + "=") != -1) {
			id += 1;
		}
		columns[index].update({'id': id});
	};
	
	this.addColumn = function (attrs) {
		var column = new Column();
		column.update(attrs);
		columns.push(column);
		checkColumnId(columns.length - 1);
		refreshIndices();
	};
	
	this.deleteColumn = function (index) {
		columns.splice(index, 1);
		refreshIndices();
	};
	
	this.updateColumn = function (index, attrs) {
		if ((index < columns.length) && (index >= 0)) {
			var column = columns[index];
			column.update(attrs);
			checkColumnId(index);

			if (index != attrs.index) {
				//теперь пересортируем по index
				var currentColumn = columns.splice(index, 1);
				if (attrs.index == 0)
					columns.unshift(currentColumn[0]);
				else {
					var part1index = attrs.index;
					var part1 = columns.slice(0, part1index);
					var part2 = columns.slice(part1index);
					columns = part1.concat(column, part2);
					refreshIndices();
				}
			}
		}
	};
	
	this.switchColumns = function (index1, index2) {
		var deleted = columns.splice(index1, 1, columns[index2]);
		columns.splice(index2, 1, deleted[0]);
		refreshIndices();
	};

	function getColumnsStreamStart () {
		var searchStr = "{\"Fixed\",\r\n";
		var pos = stream.search(searchStr);
		if (pos != -1)
			return pos + searchStr.length;
		else
			return pos;
	};
	
	this.hasColumnsContainer = function () {
		if (getColumnsStreamStart() == -1)
			return 0;
		else
			return 1;
	};
	
	function getControlsStreamStart () {
		var searchStr = "{\"Controls\",";
		var pos = stream.search(searchStr);
		if (pos != -1)
			return pos + searchStr.length;
		else {
			searchStr = "{\"Controls\"";
			pos = stream.search(searchStr);
			return pos + searchStr.length;
		}
	};
	
	this.getColumns = function () {
		var columns = [];
		var columnsStartPos = getColumnsStreamStart();
		if (columnsStartPos != -1) {
			var string = stream.substring(columnsStartPos, stream.search(/}},\r\n{\"Controls\"/g));
			string = string.replace(/(}),\r\n/g, "$1\\r\\n");
			var columnsStringArray = string.split("\\r\\n");
			for (var i = 0; i < columnsStringArray.length; i++) {
				//Message(columnsStringArray[i]);
				var column = new Column(columnsStringArray[i], i);
				columns.push(column);
			}
		}
		return columns;
	};
	
	this.getColumnByIndex = function (index) {
		return columns[index];
	};
	
	this.updateColumnsStream = function () {
		var columnsStartPos = getColumnsStreamStart();
		if (columnsStartPos != -1) {
			var newStream = stream.substring(0, columnsStartPos);
			var columnsStringArray = [];
			for (var i  = 0; i < columns.length; i++) {
				columnsStringArray.push(columns[i].Src());
			}
			newStream += columnsStringArray.join(",\r\n") + stream.substring(stream.search(/}},\r\n{\"Controls\"/g));
			stream = newStream; 
		}
	};
	
	this.updateStream = function () {
		//Message(unwrapString(stream));
		frm.Stream = unwrapString(stream);
	};
	
	this.readStream = function () {
		stream = wrapString(frm.Stream);
		//Message(stream);
 		columns = this.getColumns();
	};
	
	this.readStream();
};

function docIsOpen (_) {
	try {
		var t = doc.IsOpen;
	}
	catch (e) {
		CommonScripts.MsgBox("Редактируемый документ уже закрыт!");
		closeForm();
		return 0;
	}
	return 1;
}

function addColumn (document) {
	if (!docIsOpen()) return;
	var ds = new DialogStream(doc.Page(0));
	if (ds.hasColumnsContainer()) {
		var htmlEditor = new HTMLEditor(document, ds);
		ds.addColumn(htmlEditor.getColumnAttrs());
		ds.updateColumnsStream();
		ds.updateStream();
		htmlEditor.reload(ds.Columns());

		document.editor.selectedColumnIndex.value = ds.Columns().length;
		selectColumn(document);
	}
};

function insertColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		if (ds.hasColumnsContainer()) {
			var htmlEditor = new HTMLEditor(document, ds);
			ds.addColumn(htmlEditor.getColumnAttrs());
			ds.updateColumn(ds.Columns().length-1, {'index': index});
			ds.updateColumnsStream();
			ds.updateStream();
			htmlEditor.reload(ds.Columns());
		

			document.editor.selectedColumnIndex.value = index+1;
			selectColumn(document);
		}
	}
	else {
		addColumn(document);
	}
};

function deleteColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		if (ds.Columns().length < 2) return;
		
		ds.deleteColumn(index);
		ds.updateColumnsStream();
		ds.updateStream();
		var htmlEditor = new HTMLEditor(document, ds);
		htmlEditor.reload(ds.Columns());
		
		document.editor.selectedColumnIndex.value = Math.min(index+1, ds.Columns().length);
		selectColumn(document);
	}
};

function updateColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		var ds = new DialogStream(doc.Page(0));
		if (ds.hasColumnsContainer()) {
			var htmlEditor = new HTMLEditor(document, ds);
			var attrs = htmlEditor.getColumnAttrs();
			if (attrs.index > -1 && attrs.index < ds.Columns().length) {
				ds.updateColumn(htmlEditor.getSelectedColumnIndex(), attrs);
				ds.updateColumnsStream();
				ds.updateStream();
				htmlEditor.reload(ds.Columns());
				document.editor.selectedColumnIndex.value = attrs.index + 1;
				selectColumn(document);
			}
		}
	}
};

function moveColumnUp (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		newIndex = index-1;
		if (newIndex < 0)
			newIndex = ds.Columns().length - 1;

		if (ds.hasColumnsContainer()) {
			ds.switchColumns(index, newIndex);
			ds.updateColumnsStream();
			ds.updateStream();

			var htmlEditor = new HTMLEditor(document, ds);
			htmlEditor.reload(ds.Columns());
			
			document.editor.selectedColumnIndex.value = newIndex+1;
			selectColumn(document);
		}
	}
};

function moveColumnDown (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		var newIndex = index-0+1;
		if (newIndex > ds.Columns().length - 1)
			newIndex = 0;
		if (ds.hasColumnsContainer()) {
			ds.switchColumns(index, newIndex);
			ds.updateColumnsStream();
			ds.updateStream();
			
			var htmlEditor = new HTMLEditor(document, ds);
			htmlEditor.reload(ds.Columns());
			
			document.editor.selectedColumnIndex.value = newIndex+1;
			selectColumn(document);
		}
	}
};

function moveColumnFirst (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		if (ds.hasColumnsContainer()) {
			if (ds.Columns().length > 1 && index != 0) {
				var htmlEditor = new HTMLEditor(document, ds);
				ds.updateColumn(index, {'index': 0});
				ds.updateColumnsStream();
				ds.updateStream();
				htmlEditor.reload(ds.Columns());
				document.editor.selectedColumnIndex.value = 1;
				selectColumn(document);
			}
		}
	}
};

function moveColumnLast (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		index--;
		var ds = new DialogStream(doc.Page(0));
		var newIndex = ds.Columns().length - 1;
		if (ds.Columns().length > 1 && index != newIndex) {
			var htmlEditor = new HTMLEditor(document, ds);
			ds.updateColumn(index, {'index': newIndex});
			ds.updateColumnsStream();
			ds.updateStream();
			htmlEditor.reload(ds.Columns());
			document.editor.selectedColumnIndex.value = newIndex + 1;
			selectColumn(document);
		}
	}
};

function resetColumnEdit (document) {
};

function selectColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0) {
		var ds = new DialogStream(doc.Page(0));
		var column = ds.getColumnByIndex(index-1);

		if (column) {
			var htmlEditor = new HTMLEditor(document, ds);
			htmlEditor.selectColumn(column);
		}
	}
};

function selectNextColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0)
		currentIndex = index - 1;
	else
		currentIndex = -1;

	var ds = new DialogStream(doc.Page(0));
	var columnsCount = ds.Columns().length;
	if (columnsCount > 0) {
		if (currentIndex == columnsCount-1)
			index = 0;
		else
			index = currentIndex + 1;

		var column = ds.getColumnByIndex(index);
		var htmlEditor = new HTMLEditor(document, ds);
		htmlEditor.selectColumn(column);
	}
};

function selectPreviousColumn (document) {
	if (!docIsOpen()) return;
	var index = document.editor.selectedColumnIndex.value;
	if (index > 0)
		currentIndex = index - 1;
	else
		currentIndex = -1;

	var ds = new DialogStream(doc.Page(0));
	var columnsCount = ds.Columns().length;
	if (columnsCount > 0) {
		if (currentIndex == 0)
			index = columnsCount - 1;
		else if (currentIndex == -1)
			index = 0;
		else
			index = currentIndex - 1;

		var column = ds.getColumnByIndex(index);
		var htmlEditor = new HTMLEditor(document, ds);
		htmlEditor.selectColumn(column);
	}
};

function selectFirstColumn (document) {
	if (!docIsOpen()) return;
	var ds = new DialogStream(doc.Page(0));
	var columnsCount = ds.Columns().length;
	if (columnsCount > 0) {
		var column = ds.getColumnByIndex(0);
		var htmlEditor = new HTMLEditor(document, ds);
		htmlEditor.selectColumn(column);
	}
};

function selectLastColumn (document) {
	if (!docIsOpen()) return;
	var ds = new DialogStream(doc.Page(0));
	var columnsCount = ds.Columns().length;
	if (columnsCount > 0) {
		var column = ds.getColumnByIndex(columnsCount-1);
		var htmlEditor = new HTMLEditor(document, ds);
		htmlEditor.selectColumn(column);
	}
};

function HTMLEditor (document, ds) {
	function getColumnObject(index) {
		return ds.getColumnByIndex(index);
	};

	this.onClick = function (e) {
		var el = e.srcElement;
		if (el.tagName == "TD") {
			el = el.parentNode;
		}
		else if (el.tagName != "TR")
			return;

		_selectColumn( getColumnObject(el.getAttribute("index")) );
	};

	this.getSelectedColumnIndex = function () {
		return document.editor.selectedColumnIndex.value-1;
	};

	this.addColumn = function (column) {
		var tr = document.getElementById("columnsbody").appendChild(document.createElement("tr"));
		tr.setAttribute("index", column.attrs.index);
		tr.attachEvent("onclick", this.onClick);
		var td = tr.appendChild(document.createElement("td"));
		td.appendChild(document.createTextNode(column.attrs.index+1));
		td = tr.appendChild(document.createElement("td"));
		td.appendChild(document.createTextNode(column.attrs.id));
		td = tr.appendChild(document.createElement("td"));
		td.appendChild(document.createTextNode(column.attrs.caption));
		td = tr.appendChild(document.createElement("td"));
		td.appendChild(document.createTextNode(column.attrs.formula));
	};
	
	function _selectColumn (column) {
		var t = document.getElementById("columnsbody");
		var rows = t.getElementsByTagName("tr");
		for (var i = 0; i < rows.length; i++) {
			if (i == column.attrs.index)
				rows[i].className = "selected";
			else
				rows[i].className = "";
		}
		
		document.editor.index.value = column.attrs.index+1;
		document.editor.id.value = column.attrs.id;
		document.editor.id.setAttribute("disabled", 1-column.isLabel());
		document.editor.caption.value = column.attrs.caption;
		document.editor.formula.value = column.attrs.formula;
		document.editor.help.value = column.attrs.help;
		document.editor.tip.value = column.attrs.tip;
		document.editor.width.value = column.attrs.width;
		document.editor.location[column.attrs.location-1].checked= true;
		document.editor.invisible.checked = column.attrs.invisible;
		
		document.editor.selectedColumnIndex.value = column.attrs.index+1;

		var div = document.getElementById("columnsTableContainer");
		var a = 18*(document.editor.selectedColumnIndex.value-6);
		div.scrollTop = a;
		//document.parentWindow.scrollBy(20);
	};

	this.selectColumn = function (column) {
		_selectColumn(column);
	};
	
	function getRadioValue (radioGroup) {
		for (var i = 0; i < radioGroup.length; i++) {
			if (radioGroup[i].checked == true) 
				return radioGroup[i].value;
		}
	};
	
	this.getColumnAttrs = function () {
		return {index: document.editor.index.value-1
				, id: document.editor.id.value
				, caption: document.editor.caption.value
				, formula: document.editor.formula.value
				, help: document.editor.help.value
				, tip: document.editor.tip.value
				, width: document.editor.width.value
				, location: getRadioValue(document.editor.location)
				, invisible: document.editor.invisible.checked
				};
	};
	
	this.reload = function (columns) {
		var rowsContainer = document.getElementById("columnsbody");
		while (rowsContainer.hasChildNodes()) {
			rowsContainer.removeChild(rowsContainer.firstChild);
		}
		for (var i = 0; i < columns.length; i++) {
			this.addColumn(columns[i]);
		}
	};
	
};

function onLoad (document) {
	var ds = new DialogStream(doc.Page(0));
	var columns = ds.getColumns();
	var htmlEditor = new HTMLEditor(document, ds);
	htmlEditor.reload(columns);
};

function reloadSelf (_) {
	Scripts.Reload(SelfScript.Name)
};

function onUnload (_) {
	reloadSelf();
};

function checkWindowType (_) {
	var w = Windows.ActiveWnd;
	if (w) {
		var d = w.Document;
		if (d) {
			if (d.Type == 1) {
				if (d == docWorkBook) {
					if (d.Kind == "Document") {
						if (Metadata.FindObject(d.ID).Childs(1).Count > 0) {
							return d;
						}
					}
					else if (d.Kind == "SubList" || d.Kind == "Journal") {
						return d;
					}
				}
			}
		}
	}
	return 0;
};

var doc;

function openEditor() {
	doc = checkWindowType();
	if (doc) {
		var formPath = BinDir + "config\\html\\FormEditor.html";
		var handlers = {onLoad: onLoad, selectColumn: selectColumn
			, selectNextColumn: selectNextColumn, selectPreviousColumn: selectPreviousColumn
			, selectFirstColumn: selectFirstColumn, selectLastColumn: selectLastColumn
			, addColumn: addColumn, insertColumn: insertColumn
			, deleteColumn: deleteColumn, moveColumnUp: moveColumnUp, moveColumnDown: moveColumnDown
			, moveColumnFirst: moveColumnFirst, moveColumnLast: moveColumnLast
			, updateColumn: updateColumn, resetColumnEdit: resetColumnEdit, closeForm: closeForm, onUnload: onUnload};
		openFormWindow("Редактор колонок", formPath, handlers);
	}
};

function handleDCEvent(handlers) {
	try {		
		htmlWnd.Document.Script.setHandlers(handlers);
	}
	catch (e) {
		Message("Ошибка инициализации окна редактора", mRedErr);
		Message(e.description, mRedErr);
	}
};

var loaded = 0;
function openFormWindow(formCaption, formPath, handlers) {
	var wnd = null;
	if (wnd = CommonScripts.FindWindow(formCaption)) {
		Windows.ActiveWnd = wnd;
	}
	else {
		wnd = OpenOleForm("Shell.Explorer", formCaption);
	}
	if (!loaded) {
		SelfScript.AddNamedItem("htmlWnd", wnd, false);
		eval('function htmlWnd::DocumentComplete(d,u){ return handleDCEvent(handlers)}');
		htmlWnd.Navigate2(formPath);
	}
	else
		htmlWnd.Document.Script.onLoad();
	loaded = 1;
};

function closeForm(_) {
	Windows.ActiveWnd.Close();
	reloadSelf();
};

function init(_) // Фиктивный параметр, чтобы процедура не попадала в макросы
{
	try {
		var ocs = new ActiveXObject("OpenConf.CommonServices");
		ocs.SetConfig(Configurator);
		SelfScript.AddNamedItem("CommonScripts", ocs, false);
	}
	catch (e) {
		Message("Не могу создать объект OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("Скрипт " + SelfScript.Name + " не загружен", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}	
};

init(0);