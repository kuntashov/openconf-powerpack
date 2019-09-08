$NAME Шаблоны форм
//$Id: Шаблоны\040форм.js,v 1.6 2006/01/08 12:55:26 alest Exp $
//
//Автор Антипин Алексей aka alest (j_alesant@mail.ru)
//
//распространяется на условиях GPL

// Работает так: выгружаем всю форму или выделенные элементы, а в другой форме можем загрузить
//Шаблоны хранятся в xml-файлах: в папке базы-локальные, в папке bin/config- глобальные
//Шаблоны могут быть созданы из формы в целом и из отдельных (выделенных) контролов
// в шаблоне хранится поток формы и модуль
// модуль контрола добавляется вручную
// !!!переводы строк заменяются на строку "\r\n", соответственно могут быть случаи, когда работать не будет 
//     (если такая строка будет в форме или в модуле)
var DBPath, GlobalPath, glPrefix;

function Settings (_) {
	var path = "";
	var xml = new ActiveXObject("MSXML2.DOMDocument");
	var root;

	this.load = function (aPath) {
		path = aPath;
		xml.async = false;
		xml.load(path);
		
		if (xml.parseError.errorCode != 0) {
			//создадим файл
			xml.appendChild(xml.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'"));
			xml.appendChild(xml.createComment("Шаблоны форм и элементов диалога"));
			root = xml.appendChild(xml.createElement("Шаблоны"));
			xml.save(path);
		}
		else {
			root = xml.documentElement;
		}
	}
	
	this.writeParam = function (containerId, templateId, paramId, paramValue) {
		var pattern = "//"+containerId+"[@Имя='"+templateId+"']";
		var node = root.selectSingleNode(pattern);
		if (!node) {
			node = root.appendChild(xml.createElement(containerId));
			node.setAttribute("Имя", templateId);
		}
		node.setAttribute(paramId, paramValue.replace(/\r\n/g, "\\r\\n"));
		/*var paramNode = node.selectSingleNode("./"+paramId);
		if (!paramNode) paramNode = node.appendChild(xml.createElement(paramId));
		if (paramNode.hasChildNodes) paramNode.removeChild(paramNode.firstChild);
		paramNode.appendChild(xml.createTextNode(paramValue.replace(/\r\n/g, "\\r\\n")));*/
	}
	
	this.readParam = function (containerId, templateId, paramId) {
		var paramValue = "";
		var pattern = "//"+containerId+"[@Имя='"+templateId+"']";
		var node = root.selectSingleNode(pattern);
		if (node) {
			paramValue = node.getAttribute(paramId);
//			var paramNode = node.selectSingleNode("./"+paramId);
//			if (paramNode.hasChildNodes) paramValue = paramNode.firstChild.text;
		}
		if (paramValue) paramValue = paramValue.replace(/\\r\\n/g, "\r\n");
		return paramValue;
	}

	this.getFormsDict = function (containerId, dict, prefix) {
		var forms = root.getElementsByTagName(containerId);
		var i = 0;
		while (i < forms.length) {
			var formName = forms.item(i).getAttribute("Имя");
			dict.Add(prefix+formName, prefix+formName);
			i++;
		}
	}

	this.writeFile = function (path) {
		xml.save(path);
	}
	
}

function getFormTemplatesDict (settings, containerId, bNewMenu) {
	var dict = new ActiveXObject("Scripting.Dictionary");
	if (bNewMenu) {
		dict.Add("<Создать локальный шаблон>", "<Создать локальный шаблон>");
		dict.Add("<Создать глобальный шаблон>", "<Создать глобальный шаблон>");
	}
	settings.load(DBPath);
	settings.getFormsDict(containerId, dict, "");
	settings.load(GlobalPath);
	settings.getFormsDict(containerId, dict, glPrefix);
	return dict;
}

function checkWindowType (_) {
	var w = Windows.ActiveWnd;
	if (w) {
		var d = w.Document;
		if (d) {
			if (d.Type == 1) {
				if (d == docWorkBook) return d;
			}
		}
	}
	return 0;
}

function createFormTemplate () {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "Форма", 1);

	var path = DBPath;
	var templateId = CommonScripts.SelectValue(templatesDict, "Выберите шаблон");
	if (!templateId) return;
	else if (templateId == "<Создать локальный шаблон>") {
		templateId = CommonScripts.InputBox("Введите наименование нового локального шаблона");
	}
	else if (templateId == "<Создать глобальный шаблон>") {
		templateId = CommonScripts.InputBox("Введите наименование нового глобального шаблона");
		path = GlobalPath;
	}
	else if (templateId.indexOf(glPrefix) != -1) {
		templateId = templateId.substr(glPrefix.length);
		path = GlobalPath;
	}

	if (!templateId) return;
	var d = checkWindowType(0);
	if (d) {
		var ds = new DialogStream(d.Page(0));

		settings.load(path);
		settings.writeParam("Форма", templateId, "Поток", ds.getStream());
		settings.writeParam("Форма", templateId, "Типы", ds.getTypesString(ds.getControlsStringArray()));
		settings.writeParam("Форма", templateId, "Модуль", d.Page(1).Text);
		settings.writeFile(path);
	}
}

function createControlsTemplate () {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "Элемент", 1);

	var path = DBPath;
	var templateId = CommonScripts.SelectValue(templatesDict, "Выберите шаблон");
	if (!templateId) return;
	else if (templateId == "<Создать локальный шаблон>") {
		templateId = CommonScripts.InputBox("Введите наименование нового локального шаблона");
		path = DBPath;
	}
	else if (templateId == "<Создать глобальный шаблон>") {
		templateId = CommonScripts.InputBox("Введите наименование нового глобального шаблона");
		path = GlobalPath;
	}
	else if (templateId.indexOf(glPrefix) != -1) {
		templateId = templateId.substr(glPrefix.length);
		path = GlobalPath;
	}

	if (!templateId) return;
	var d = checkWindowType(0);
	if (d) {
		var ds = new DialogStream(d.Page(0));
		
		settings.load(path);
		settings.writeParam("Элемент", templateId, "Поток", ds.replaceLayerName(ds.getSelectedControlsString(), ds.getActiveLayerName(), "Основной"));
		settings.writeParam("Элемент", templateId, "Типы", ds.getTypesString(ds.getSelectedControlsStringArray()));
		settings.writeFile(path);
	}
}

function DialogStream (frm) {
	/*формат:
	с 5-ой строки описание слоев:
		{"1",
		{"Основной","1"},
		{"Слой1","0"}},"1","1"}},
		номер активного слоя (с нуля)
		название слоя, флаг видимости
		

	раздел Controls:
	элементы идут в обратном порядке (по отношению к порядку перехода по Tab)
	Control:
	1. Видимый текст
	2. Тип : BUTTON, CHECKBOX, RADIO, STATIC (текст), 1CGROUPBOX, LISTBOX (список), COMBOBOX (поле со списком), 
		BMASKED (дата, агрегатного типа)
		1CEDIT (неопр, строка, число)
	3. Код типа:
		BUTTON = 1342177291
		CHECKBOX = 1342177283
		RADIO = 1342177289
		STATIC = 1342177792
		1CGROUPBOX = 1342177287
		LISTBOX = 1352663297
		COMBOBOX = 1352663107
		TABLE = 1352663040
		BMASKED = 1350565888
		1CEDIT = 1350565888
	4. X
	5. Y
	6. Длина
	7. Высота
	8. 0
	9. 0
	10. Порядок следования: чем больше номер, тем раньше строит в порядке следования (4153 и больше)
	11. пусто
	12. Формула
	13. Идентификатор
	14. 0 или -1
	15. B, U, T, N
	16. длина (для числа, строки)
	17. точность
	18. Числовой код 1с-типа контрола (Справочник.Контрагенты ... )
	19. 0
	20. число
	21. пусто
	22. описание
	23. подсказка
	24. 0
	25. -11 или -13
	26. 0
	27. 0
	28. 0
	29-37. число
	38. шрифт - MS Sans Serif
	39. -1
	40. -1
	41. 0
	42. название слоя  - "Слой2"
	43. {""0"",""0""}
	
	
	
	*/

	var stream = frm.Stream;	

	var layers = [];
	var activeLayer;
	
	var controls = [];
	
	this.setStream = function (newStream) {
		stream = newStream;
		frm.Stream = stream;
	}
	
	this.getStream = function () {
		stream = frm.stream;
		return stream;
	}
	
	this.getActiveLayerName = function () {
		return frm.ActiveLayer;
	}
	
	this.getLayersCollection = function () {
		for (var i = 0; i < frm.LayerCount; i++) {
			layers.push(frm.LayerName(i));
		}
	}
	
	this.replaceLayerName = function (str, currLayerName, newLayerName) { //меняем на Основной
		re = new RegExp();
		re.compile("\"" + currLayerName + "\",\"{\"\"0\"\",\"\"0\"\"}\"}", "g");
		str = str.replace(re, "\"" + newLayerName + "\",\"{\"\"0\"\",\"\"0\"\"}\"}");
		return str;
	}

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
	}
	
	function unwrapString (string) {
		string = string.replace(/k2av/g, "\"");
		string = string.replace(/s5ace/g, "\r\n");

		return string;
	}
	
	this.getTypesString = function (stringArr) {
		var types = [];

		for (var i = 0; i < stringArr.length; i++) {
			var string = wrapString(stringArr[i]);
			string = string.replace(/("),(")/g, "$1\r\n$2");
			var controlParams = string.split("\r\n");
			var arr = [];
			var typeFullName = "";
			var typeId = controlParams[17].substr(1, controlParams[17].length-2);
			if (typeId != "0") {
				typeFullName = Metadata.FindObject(typeId).FullName;
			}
			arr.push(typeFullName);
			arr.push(controlParams[14]);
			arr.push(controlParams[15]);
			arr.push(controlParams[16]);
			arr.push(controlParams[17]);

			types.push(arr.join(","));
		}
		return types.join("\\r\\n");
	}
	
	function fillMDhash (mdObjs, hash, type) {
		for (var i = 0; i < mdObjs(type).Count; i++) {
			hash[mdObjs(type)(i).FullName] = mdObjs(type)(i).ID;
		}
	}

	this.checkTypes = function (controlsString, typesString) {	//при загрузке проверим типы
		if (!typesString) return controlsString;
		
		var controlsLines = controlsString.split("},\r\n");
		var typesLines = typesString.split("\r\n");

		var mdObjsHash = {};
		var mdObjs = MetaData.TaskDef.Childs;
		fillMDhash(mdObjs, mdObjsHash, "Справочник");
		fillMDhash(mdObjs, mdObjsHash, "Документ");
		fillMDhash(mdObjs, mdObjsHash, "Перечисление");
		//fillMDhash(mdObjs, mdObjsHash, "Счет");
				
		for (var i = 0; i < controlsLines.length; i++) {
			var types = typesLines[i].split(",");
			var typeFullName = types.shift();
			var oldTypeString = types.join(",");

			if (!mdObjsHash[typeFullName]) {
				Message("Не найден тип " + typeFullName + "!");
				types[4-1] = "\"0\"";
			}
			else {
				types[4-1] = "\""+ mdObjsHash[typeFullName] +"\"";
			}
			var newTypeString = types.join(",");
			var re = new RegExp();
			re.compile(oldTypeString, "g");
			controlsLines[i] = controlsLines[i].replace(re, newTypeString);
		}
		
		return controlsLines.join("},\r\n");
	}
	
	this.prepareLoadForm = function (formString, typesString) {
		stream = formString;
		var newControlsString = this.checkTypes(this.replaceLayerName(this.getControlsString(), "Основной", frm.ActiveLayer), typesString);
		var searchStr1 = "{\"Controls\",\r\n";
		var searchStr2 = "},\r\n{\"Cnt_Ver\",\"10001\"}}";
		stream = stream.substring(0, stream.indexOf(searchStr1)+searchStr1.length) + newControlsString + stream.substr(stream.indexOf(searchStr2));
		return stream;
	}

	this.prepareLoadControls = function (controlsString, typesString) {
		return this.checkTypes(this.replaceLayerName(controlsString, "Основной", frm.ActiveLayer), typesString);
	}
		
	this.addControls = function (string) {
		var searchStr = "{\"Controls\",\r\n";
		if (frm.stream.search(searchStr) == -1) {
			stream = stream.replace(/{"Controls"/, searchStr);
		}
		else {
			string += ",\r\n";
		}
		var controlsStartPos = stream.indexOf(searchStr);
		var newStream = stream.substring(0, controlsStartPos + searchStr.length) + string 
			+ stream.substring(controlsStartPos + searchStr.length);
		this.setStream(newStream);
	}
	
	this.getControlsString = function () {
		var string = "";
		var searchStr1 = "{\"Controls\",\r\n";
		var searchStr2 = "},\r\n{\"Cnt_Ver\",\"10001\"}}";
		var controlsStartPos = stream.indexOf(searchStr1);
		
		if (controlsStartPos != -1) {
			string = stream.substring(controlsStartPos + searchStr1.length, stream.indexOf(searchStr2));
		}
		return string;
	}

	this.getControlsStringArray = function () {
		var string = this.getControlsString();
		string = string.replace(/(}),\r\n/g, "$1\\r\\n");
		return string.split("\\r\\n");
	}
	
	this.getSelectedControlsString = function () {
		return this.getSelectedControlsStringArray().join(",\r\n");
	}
	
	this.getSelectedControlsStringArray = function () {
		var selectedControlsStringArray = [];
		if (frm.Selection) {
			var controlsStringArray = this.getControlsStringArray();
			controlsStringArray.reverse();
			//таблица тоже участвует в порядке обхода
			//{"Browser","8","0", - второе число - порядковый номер
			var re = /{\"Browser\",\"([^"]*)"/g;
			var matches = re.exec(frm.Stream);
			var browserNum = -1;
			if (matches) {
				browserNum = matches[1];
				var t = controlsStringArray.slice(0, browserNum);
				t.push("!Browser!");
				controlsStringArray = t.concat(controlsStringArray.slice(browserNum));
			}
			var selControls = frm.Selection.split(",");
			for (var i = 0; i < selControls.length; i++) {
				if (selControls[i] == browserNum) continue;
				selectedControlsStringArray.push(controlsStringArray[selControls[i]]);
			}
		}
		return selectedControlsStringArray;
	}
}

function loadForm() {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "Форма", 0);
	var templateId = CommonScripts.SelectValue(templatesDict, "Выберите шаблон");
	if (!templateId) return;
	
	if (templateId.substr(0, glPrefix.length) == glPrefix) {
		templateId = templateId.substr(glPrefix.length);
		settings.load(GlobalPath);
	}
	else {
		settings.load(DBPath);
	}

	var d = checkWindowType(0);
	if (d) {
		ds = new DialogStream(d.Page(0));
		ds.setStream(ds.prepareLoadForm(settings.readParam("Форма", templateId, "Поток"), settings.readParam("Форма", templateId, "Типы")));
		d.Page(1).text = settings.readParam("Форма", templateId, "Модуль");
	}
}

function loadControls () {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "Элемент", 0);
	var templateId = CommonScripts.SelectValue(templatesDict, "Выберите шаблон");
	if (!templateId) return;
	
	if (templateId.substr(0, glPrefix.length) == glPrefix) {
		templateId = templateId.substr(glPrefix.length);
		settings.load(GlobalPath);
	}
	else {
		settings.load(DBPath);
	}

	var d = checkWindowType(0);
	if (d) {
		var ds = new DialogStream(d.Page(0));
		ds.addControls(ds.prepareLoadControls(settings.readParam("Элемент", templateId, "Поток"), settings.readParam("Форма", templateId, "Типы")));
		var moduleText = settings.readParam("Элемент", templateId, "Модуль");
		if (moduleText) d.Page(1).text += moduleText;
	}
}

function printDialogStream () {
	var d = checkWindowType(0);
	if (d) Message(d.Page(0).Stream);
}

//заменить строку в идентификаторах/заголовках
// 
function replaceString () {
	var d = checkWindowType(0);
	if (d) {
		var actionType = CommonScripts.SelectValue("идентификаторах\r\nзаголовках\r\nформулах\r\n", "Замена в ");
		if (!actionType) return;
		
		var searchString = CommonScripts.InputBox("Регэксп поиска");
		if (searchString) {
			var frm = d.Page(0);
			var replaceString = CommonScripts.InputBox("Регэксп замены");
			var searchRe = new RegExp();
			searchRe.compile(searchString, "i");

			var propId;
			if (actionType == "идентификаторах") {
				propId = cpStrId;
			}
			else if (actionType == "заголовках") {
				propId = cpTitle;
			}
			else if (actionType == "формулах") {
				propId = cpFormul;
			}

			for (var i = 0; i < frm.ctrlCount; i++) {
				frm.ctrlProp(i, propId) = frm.ctrlProp(i, propId).replace(searchRe, replaceString);
			}
		}
	}
}

function init(_) // Фиктивный параметр, чтобы процедура не попадала в макросы
{
	try {
		var ocs = new ActiveXObject("OpenConf.CommonServices");
		ocs.SetConfig(Configurator);
		SelfScript.AddNamedItem("CommonScripts", ocs, false);
		GlobalPath = CommonScripts.FSO.BuildPath(BinDir, "config\\Шаблоны форм.xml");
		DBPath = CommonScripts.FSO.BuildPath(IBDir, "Шаблоны форм.xml");
		glPrefix = "(гл)_";
	}
	catch (e) {
		Message("Не могу создать объект OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("Скрипт " + SelfScript.Name + " не загружен", mInformation);
		Scripts.UnLoad(SelfScript.Name);
	}
}

init(0);
