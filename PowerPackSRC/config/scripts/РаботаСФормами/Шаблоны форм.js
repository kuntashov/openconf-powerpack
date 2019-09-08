$NAME ������� ����
//$Id: �������\040����.js,v 1.6 2006/01/08 12:55:26 alest Exp $
//
//����� ������� ������� aka alest (j_alesant@mail.ru)
//
//���������������� �� �������� GPL

// �������� ���: ��������� ��� ����� ��� ���������� ��������, � � ������ ����� ����� ���������
//������� �������� � xml-������: � ����� ����-���������, � ����� bin/config- ����������
//������� ����� ���� ������� �� ����� � ����� � �� ��������� (����������) ���������
// � ������� �������� ����� ����� � ������
// ������ �������� ����������� �������
// !!!�������� ����� ���������� �� ������ "\r\n", �������������� ����� ���� ������, ����� �������� �� ����� 
//     (���� ����� ������ ����� � ����� ��� � ������)
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
			//�������� ����
			xml.appendChild(xml.createProcessingInstruction("xml", "version='1.0' encoding='windows-1251'"));
			xml.appendChild(xml.createComment("������� ���� � ��������� �������"));
			root = xml.appendChild(xml.createElement("�������"));
			xml.save(path);
		}
		else {
			root = xml.documentElement;
		}
	}
	
	this.writeParam = function (containerId, templateId, paramId, paramValue) {
		var pattern = "//"+containerId+"[@���='"+templateId+"']";
		var node = root.selectSingleNode(pattern);
		if (!node) {
			node = root.appendChild(xml.createElement(containerId));
			node.setAttribute("���", templateId);
		}
		node.setAttribute(paramId, paramValue.replace(/\r\n/g, "\\r\\n"));
		/*var paramNode = node.selectSingleNode("./"+paramId);
		if (!paramNode) paramNode = node.appendChild(xml.createElement(paramId));
		if (paramNode.hasChildNodes) paramNode.removeChild(paramNode.firstChild);
		paramNode.appendChild(xml.createTextNode(paramValue.replace(/\r\n/g, "\\r\\n")));*/
	}
	
	this.readParam = function (containerId, templateId, paramId) {
		var paramValue = "";
		var pattern = "//"+containerId+"[@���='"+templateId+"']";
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
			var formName = forms.item(i).getAttribute("���");
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
		dict.Add("<������� ��������� ������>", "<������� ��������� ������>");
		dict.Add("<������� ���������� ������>", "<������� ���������� ������>");
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
	var templatesDict = getFormTemplatesDict(settings, "�����", 1);

	var path = DBPath;
	var templateId = CommonScripts.SelectValue(templatesDict, "�������� ������");
	if (!templateId) return;
	else if (templateId == "<������� ��������� ������>") {
		templateId = CommonScripts.InputBox("������� ������������ ������ ���������� �������");
	}
	else if (templateId == "<������� ���������� ������>") {
		templateId = CommonScripts.InputBox("������� ������������ ������ ����������� �������");
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
		settings.writeParam("�����", templateId, "�����", ds.getStream());
		settings.writeParam("�����", templateId, "����", ds.getTypesString(ds.getControlsStringArray()));
		settings.writeParam("�����", templateId, "������", d.Page(1).Text);
		settings.writeFile(path);
	}
}

function createControlsTemplate () {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "�������", 1);

	var path = DBPath;
	var templateId = CommonScripts.SelectValue(templatesDict, "�������� ������");
	if (!templateId) return;
	else if (templateId == "<������� ��������� ������>") {
		templateId = CommonScripts.InputBox("������� ������������ ������ ���������� �������");
		path = DBPath;
	}
	else if (templateId == "<������� ���������� ������>") {
		templateId = CommonScripts.InputBox("������� ������������ ������ ����������� �������");
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
		settings.writeParam("�������", templateId, "�����", ds.replaceLayerName(ds.getSelectedControlsString(), ds.getActiveLayerName(), "��������"));
		settings.writeParam("�������", templateId, "����", ds.getTypesString(ds.getSelectedControlsStringArray()));
		settings.writeFile(path);
	}
}

function DialogStream (frm) {
	/*������:
	� 5-�� ������ �������� �����:
		{"1",
		{"��������","1"},
		{"����1","0"}},"1","1"}},
		����� ��������� ���� (� ����)
		�������� ����, ���� ���������
		

	������ Controls:
	�������� ���� � �������� ������� (�� ��������� � ������� �������� �� Tab)
	Control:
	1. ������� �����
	2. ��� : BUTTON, CHECKBOX, RADIO, STATIC (�����), 1CGROUPBOX, LISTBOX (������), COMBOBOX (���� �� �������), 
		BMASKED (����, ����������� ����)
		1CEDIT (�����, ������, �����)
	3. ��� ����:
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
	6. �����
	7. ������
	8. 0
	9. 0
	10. ������� ����������: ��� ������ �����, ��� ������ ������ � ������� ���������� (4153 � ������)
	11. �����
	12. �������
	13. �������������
	14. 0 ��� -1
	15. B, U, T, N
	16. ����� (��� �����, ������)
	17. ��������
	18. �������� ��� 1�-���� �������� (����������.����������� ... )
	19. 0
	20. �����
	21. �����
	22. ��������
	23. ���������
	24. 0
	25. -11 ��� -13
	26. 0
	27. 0
	28. 0
	29-37. �����
	38. ����� - MS Sans Serif
	39. -1
	40. -1
	41. 0
	42. �������� ����  - "����2"
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
	
	this.replaceLayerName = function (str, currLayerName, newLayerName) { //������ �� ��������
		re = new RegExp();
		re.compile("\"" + currLayerName + "\",\"{\"\"0\"\",\"\"0\"\"}\"}", "g");
		str = str.replace(re, "\"" + newLayerName + "\",\"{\"\"0\"\",\"\"0\"\"}\"}");
		return str;
	}

	function wrapString (string) {	//������� ������ ������
		string = string.replace(/""{([^{(),"}]+)}""/g, "sendkey|$1|"); //EmulationKey(""{RIGHT}"",0);
		string = string.replace(/(,|\()"""",/g, "$1\k2avk2av,"); // fix bug with �������.frm and others with [,""""",] inside atom #1880
								// and ��������������������("""", to #3012
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

	this.checkTypes = function (controlsString, typesString) {	//��� �������� �������� ����
		if (!typesString) return controlsString;
		
		var controlsLines = controlsString.split("},\r\n");
		var typesLines = typesString.split("\r\n");

		var mdObjsHash = {};
		var mdObjs = MetaData.TaskDef.Childs;
		fillMDhash(mdObjs, mdObjsHash, "����������");
		fillMDhash(mdObjs, mdObjsHash, "��������");
		fillMDhash(mdObjs, mdObjsHash, "������������");
		//fillMDhash(mdObjs, mdObjsHash, "����");
				
		for (var i = 0; i < controlsLines.length; i++) {
			var types = typesLines[i].split(",");
			var typeFullName = types.shift();
			var oldTypeString = types.join(",");

			if (!mdObjsHash[typeFullName]) {
				Message("�� ������ ��� " + typeFullName + "!");
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
		var newControlsString = this.checkTypes(this.replaceLayerName(this.getControlsString(), "��������", frm.ActiveLayer), typesString);
		var searchStr1 = "{\"Controls\",\r\n";
		var searchStr2 = "},\r\n{\"Cnt_Ver\",\"10001\"}}";
		stream = stream.substring(0, stream.indexOf(searchStr1)+searchStr1.length) + newControlsString + stream.substr(stream.indexOf(searchStr2));
		return stream;
	}

	this.prepareLoadControls = function (controlsString, typesString) {
		return this.checkTypes(this.replaceLayerName(controlsString, "��������", frm.ActiveLayer), typesString);
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
			//������� ���� ��������� � ������� ������
			//{"Browser","8","0", - ������ ����� - ���������� �����
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
	var templatesDict = getFormTemplatesDict(settings, "�����", 0);
	var templateId = CommonScripts.SelectValue(templatesDict, "�������� ������");
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
		ds.setStream(ds.prepareLoadForm(settings.readParam("�����", templateId, "�����"), settings.readParam("�����", templateId, "����")));
		d.Page(1).text = settings.readParam("�����", templateId, "������");
	}
}

function loadControls () {
	var settings = new Settings();
	var templatesDict = getFormTemplatesDict(settings, "�������", 0);
	var templateId = CommonScripts.SelectValue(templatesDict, "�������� ������");
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
		ds.addControls(ds.prepareLoadControls(settings.readParam("�������", templateId, "�����"), settings.readParam("�����", templateId, "����")));
		var moduleText = settings.readParam("�������", templateId, "������");
		if (moduleText) d.Page(1).text += moduleText;
	}
}

function printDialogStream () {
	var d = checkWindowType(0);
	if (d) Message(d.Page(0).Stream);
}

//�������� ������ � ���������������/����������
// 
function replaceString () {
	var d = checkWindowType(0);
	if (d) {
		var actionType = CommonScripts.SelectValue("���������������\r\n����������\r\n��������\r\n", "������ � ");
		if (!actionType) return;
		
		var searchString = CommonScripts.InputBox("������� ������");
		if (searchString) {
			var frm = d.Page(0);
			var replaceString = CommonScripts.InputBox("������� ������");
			var searchRe = new RegExp();
			searchRe.compile(searchString, "i");

			var propId;
			if (actionType == "���������������") {
				propId = cpStrId;
			}
			else if (actionType == "����������") {
				propId = cpTitle;
			}
			else if (actionType == "��������") {
				propId = cpFormul;
			}

			for (var i = 0; i < frm.ctrlCount; i++) {
				frm.ctrlProp(i, propId) = frm.ctrlProp(i, propId).replace(searchRe, replaceString);
			}
		}
	}
}

function init(_) // ��������� ��������, ����� ��������� �� �������� � �������
{
	try {
		var ocs = new ActiveXObject("OpenConf.CommonServices");
		ocs.SetConfig(Configurator);
		SelfScript.AddNamedItem("CommonScripts", ocs, false);
		GlobalPath = CommonScripts.FSO.BuildPath(BinDir, "config\\������� ����.xml");
		DBPath = CommonScripts.FSO.BuildPath(IBDir, "������� ����.xml");
		glPrefix = "(��)_";
	}
	catch (e) {
		Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("������ " + SelfScript.Name + " �� ��������", mInformation);
		Scripts.UnLoad(SelfScript.Name);
	}
}

init(0);
