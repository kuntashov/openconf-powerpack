/* ============================================================================ *\
������:  als2xml.js
������:  $Revision: 1.7 $
�����:   ��������� �������� aka a13x
E-mail:  <kuntashov@yandex.ru>
ICQ UIN: 338758861
��������: 
	������ ��� ��������� �� ������ als-������ ���������-��������� 1�:�����������
xml-������ �������� ����� ��� ��������.
\* ============================================================================ */

var OpenConf = false; // true - �������� �� � ��������� OpenConf'�, ����� - WSH

var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

/* ============================================================================ *\
   								"�������" ������
\* ============================================================================ */

// true - "�����" ����� (�� ��������� ������� ���������), false - ������� �����
var Silent = false;

function usage(_)
{
	return ( 
'��������������� als-������ � xml � ������� ������� xml2tls.exe ���������� ��������\n\
\n\
�������������:\n\
cscript.exe //nologo als2xml.js /ALS:���ALS����� [������������ ���������]\n\
\n\
���������:\n\
/ALS:<���� � ��� ALS-�����> - ������������\n\
	����/��� ��������� als-�����\n\
\n\
/XML[:<���� � ��� XML-�����>] - ��������������\n\
	���� � ��� ������������� xml-�����. �� ��������� xml-���� ��������� \n\
	� ��� �� ������, ��� � �������� als (� � ��� �� ��������), �� �\n\
	����������� "xml".\n\
\n\
/TLS[:<���� � ��� TLS-�����>] - ��������������\n\
	������������� ������� ������� xml2tls.exe ��� ���������\n\
	�� ������ ����������� xml-����� tls-����� ��� ��������.\n\
	xml2tls ������ ���������� � ����� �������� �� ��������, ����\n\
	� ����� �� ���������, ����������� � ���������� ��������� PATH.\n\
	���� ��� � ���� ����� �� ������, �� ���� ����������� � ��������\n\
	� �������� als-������ � ��� �� ������, �� � ����������� "tls".\n\
\n\
/INTS:<������� ����������> - ��������������\n\
	������������� ints-����� ��� ������� Intellisence.vbs, ��������\n\
	�� � ��������� ��������. ����� ������ ����� ��������� � �������\n\
	�����. ����� ����� ������ ���� �������������.ints, � ������� �����\n\
	�������� ����� ����� � ��� �� �������, ��� ������������ � �����������\n\
	��������� ����� ������� Intellisence.vbs.\n\
	��������. �� ������������� � �������� �������� ���������� ������������\n\
	������� <���������>\\config\\Intell, �.�. ����� ������������ �� ������\n\
	������ � als-������ � �������� ����� � ��� �� ������ �������������\n\
	progid-� �������.\n\
\n\
/NOXML - ��������������\n\
	�� ��������� xml-����. ����� ����� ������������ ��������� �\n\
	������ /TLS. ���� �� �������� ��� �����, �� ������ � tls-������\n\
	����� �������� � ��������������� xml-����.\n\
\n\
/JS[:<��� � ���� � js-�����>] - ��������������\n\
	��������� ������������� ��������� �������� ALS-����� (��� �������).\n\
	���� ��� � ���� ����� �� ������, �� ���� ����������� � ��������\n\
	� �������� als-������ � ��� �� ������, �� � ����������� "js".\n\
\n\
/PREFIX[:<�������>] - ��������������\n\
	������������ ������� <�������> ��� ��������������� �����,\n\
	����������� �� ��������������� als-�����. ���� ���� ����� ���\n\
	��������� <�������>, �� � �������� �������� ������������ ���\n\
	��������� als-����� (��� ����������).\n\
\n\
/INDENT[:<������������������>] - ��������������\n\
	������������� ������������ XML-��� ��������� (�� ��������,\n\
	��� �������� ������������ �������). �������������� ��������\n\
	�������� <������������������> ������ ������� ������� �� ������\n\
	������, �� ��������� ������������ �������� 2.\n\
\n\
/SILENT - ��������������\n\
	�� �������� ������� ���������.');
}

function xml2tls(xmlFile, tlsFile) 
{
	var xml2tls_exe = "";
	
	if (OpenConf)
		xml2tls_exe= fso.BuildPath(BinDir, "config\\system\\xml2tls\\xml2tls.exe");
	
	var cmd = 'xml2tls.exe "' + xmlFile + '" "' + tlsFile + '"';
	
	if (fso.FileExists(xml2tls_exe)) {
		cmd = cmd.replace(/^xml2tls\.exe/, '"' + xml2tls_exe + '"')
	}
	
	try {
		return (0 == wsh.Run(cmd, 7, true));
	}
	catch (e) {
		msg("������ ������� xml2tls: ��������, ������� �� ��������� � PATH.");
	}
}

function makePath(alsFile, ext)
{
	return fso.BuildPath(wsh.CurrentDirectory, fso.GetBaseName(alsFile) + '.' + ext);
}

function Run(_) 
{
	var prefix="", indent="";
	var Args = WScript.Arguments.Named;
	
	if (!Args.Count || Args.Exists("?") || Args.Exists("HELP") || Args.Exists("H")) {
		msg(usage());
		exit(1);
	}
	
	if (!Args.Exists("ALS")) {
		msg(usage());
		exit(1);
	}
		
	var alsFile, xmlFile, tlsFile, jsFile;
	var intsFolder = null;

	var noXML = Args.Exists("NOXML");
	Silent = Args.Exists("SILENT");

	alsFile = Args.Item('ALS');
	if (!fso.FileExists(alsFile)) {
		msg("��������� als-���� �� ����������!");
		exit(1);
	}

	if (Args.Exists("PREFIX")) {
		prefix = Args.Item("PREFIX");
		if (!prefix) {
			prefix = fso.GetBaseName(alsFile);
		}
		// ����� ������ ����� ���� � ���������
		prefix = makeValidId(prefix);
	}	

	if (Args.Exists("INTS")) {
		intsFolder = Args.Item("INTS");
		if (!fso.FolderExists(intsFolder)) {
			msg("��������� ������� ��� ��������� ints-������ �� ����������!");
			exit(1);
		}
	}
	
	// XXX ���� �� ����� /XML � ����� /NOXML, �� xml-���� ��������� � temp'� (?)

	if (Args.Exists("XML") && !Args.Item("XML")) {
		xmlFile = Args.Item("XML");
	}
	else {
		xmlFile = makePath(alsFile, "xml");
	}
	
	if (Args.Exists("TLS")) {
		if (Args.Item("TLS")) {
			tlsFile = Args.Item("TLS");
		}
		else {
			tlsFile = makePath(alsFile, "tls");
		}
	}
	
	if (Args.Exists("JS")) {
		if (Args.Item("JS")) {			
			jsFile = Args.Item("JS");
		}
		else {
			jsFile = makePath(alsFile, "js");
		}
	}		

	if (Args.Exists("INDENT")) {
		indentSize = Args.Item("INDENT");
		if (!indentSize) {
			indentSize = 2;
		}
		for ( ; indent.length<indentSize; indent+=" ")
			; // ���-�� ������� ��� ������ �����, �� �� �� �� ����� :-) -- a13x
	}	

	msg("������� ��������� als-�����: " + alsFile);
	var tree = parseAls(alsFile, jsFile);

	if (tree) {			

		msg("���������� ����������� ������...");
		var types = processShell(tree);

		if (!noXML||tlsFile) {
			msg("��������� XML-�����...");
			saveAsXML(types, xmlFile, prefix, indent);
		}
	
		var ret = false;
		if (tlsFile) {
			msg("���������� XML-�����...");
			ret = xml2tls(xmlFile, tlsFile);
		}

		if (noXML && fso.FileExists(xmlFile)) {
			msg("�������� ��������� ������...");
			fso.DeleteFile(xmlFile, true);
		}

		if (intsFolder) {
			msg("��������� ints-������...");
			saveAsInts(types, intsFolder, prefix);
		}

		msg("---------------------------------------");
		
		if (!noXML)			msg("XML: " + xmlFile);
		if (tlsFile && ret)	msg("TLS: "	+ tlsFile);
		if (jsFile)			msg("JS:  " + jsFile);		

	}
	else {
		exit(1);
	}
}

/* ============================================================================ *\
   								������ ��� OPENCONF
\* ============================================================================ */

function Init(_) // ��������� ��������, ����� ��������� �� �������� � �������
{
	try {
		var ocs = new ActiveXObject("OpenConf.CommonServices");
		ocs.SetConfig(Configurator);
		SelfScript.AddNamedItem("CommonScripts", ocs, false);
	}
	catch (e) {
		Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("������ " + SelfScript.Name + " �� ��������", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}	
}

function ��������������ALS() 
{
	var alsFile = CommonScripts.SelectFileForRead(BinDir + "*.als", "����� ���������-��������� (*.als)|*.als|��� �����|*");
		
	if (alsFile) {		
		var xmlFile = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName() + ".xml");
		
		msg("�������������� als-����: " + alsFile);
		var tree = parseAls(alsFile);

		if (tree) {			

			var tlsFile = alsFile.replace(/als$/, "tls");
			
			msg("���������� ����������� ������...");
			var types = processShell(tree);
			
			msg("��������� XML-�����...");
			saveAsXML(types, xmlFile);
	
			var ret = false;
			msg("���������� XML-�����...");
			ret = xml2tls(xmlFile, tlsFile);			
			
			msg("�������� ��������� ������...");
			fso.DeleteFile(xmlFile, true);

			if (ret)
				msg("������!");
			else
				msg("���� �� ������������", mRedErr);
		}
	}
}

/* ============================================================================ *\

���� ���-�� �������� ����������� � �����, �� ��� �������� ����� ������, ���� 
�������� ��� ����� ���� �������� �� ��������� ������� ��� ������� �����������
����, ��� ��� ��� ������ ��������, � ���� ����� ��������� ������ ���������������
������ ������������� � ������ ������� �������. -- a13x

ALS-���� ������������ ����� ��������� ���� ���������� �������: 

{"Shell",
 {"Folder","AST","����������","ObjectName",
    {"Item","AST","�����������","PropertyName",
		"�����������","PropertyName","������������� ����� �������� ��������"},
    {"Item","AST","���������","MethodName",
		"���������()","MethodName()","������������� ����� �������� ������"},
	...
 },
 ...
}

�������� ���� Folder ����� ����� ������ ��������� �������� ���� Folder (�� 
���������� �� ���������� ������� �����������).

�������� �������������� ���������:

��� 1. �������������� als-����� � js-������. 
--------------------------------------------

��� �������������� ���������� �������, ������������ ������� ������� � �������
mms.view.py, �������� � �������� ������� GComp (http://1c.alterplast.ru),
� ������: ���������� ���, ��� ��������� als-����� ���������� �� ����������
����������� ������� ������� � js (� ��������� - ������), ����� ������ �������� 
�������� ������ ("{" � "}") �� ���������� ("[" � "]" ��������������) � ������ 
�������� � ������������� ������� � �������������.

����� �������������� ����������� � ������� parseAls().

��� 2. ���������� ���������� ��������� � ������������ � ��� ������.
-------------------------------------------------------------------

�� ������ ����� �� ������ ��������� ��������:

	1. als-���� ����� ��������� ��������� ������������ �����������, � �� �����
		��� � xml-����� ��������� �������� ������ ������ � ���������� ����� ��������:

			 ��� (������)
			 |	
			 +---������ �/��� ��������

	2. ������������ ����� ���������� � ������� � ��������� ���� (�������) ������������
		� ���� plain-text, ������� ������������ ������� ������������, ��, � ���������,
		������ ���� �������� �� ������� � � ���� ���� ���-�� ��������

� ��� �������, ����� �������� ������� � ������� � �������� ������������� � ������� (�������)
� ������� "������" � "��������" ��������������, ��������, ����� ����� ���������� �����
�������� �� �������� � ��������� ���, � "����������" ��������, �.�. �������� ��������� 
�������� � ��������� ������������� �������.

�������������� ��������� �� ��������� �����:

1. ��� ������� ������� �����:
2. ���������� ������ � ������ �����
3. ���������� ��������� �������� �������:
	3.1 ������� ������ ����� ��������� � ����������� �� ���� ������ ���� � ������
	�������, ���� � ������ ������� ����, ���������������� ��������������� �������
	3.2. ���� ��������� ������� - ������, ��
		3.2.1 ���� ��� ������� "��������" ��� ��� ������� "������", �� ���������
		��� �������� (� ��� �������� ������ ���� "Item") � ������ ������� ��� ������� 
		������� ��������������
		3.2.2 ����� �������, ��� ��� ����� ������ � ��������� ���������� � ����� ������� 
		��������� ��������, ������� � ������ 2.

��� 3. ���������� ������������ ��������� ��������� � XML.
---------------------------------------------------------

����� ��� ���������� ������, ��� �� ������ ����������� �� ���������� ������ :-).
   
\* ============================================================================ */

/* ============================================================================ *\
                    ������� �������������� ALS-����� � JS-������
\* ============================================================================ */

/* 
 * ������� ������ ������ aka fez �� ������������ ����, ��� ����� 
 * ������������� ���� � mms.view.py -- a13x
 */
function parseAls(alsFile, jsFile) 
{	
	var tree = null;

	var out = jsFile ? fso.OpenTextFile(jsFile, 2, true) : null;
	var inp = fso.OpenTextFile(alsFile, 1);
	
	var s = inp.ReadAll();

/*

0. ���������� �����

1. ������ ����������� ������� ������� �� ���������: 
	
	1.1 ��������� ������� ������ ��������� ���������� �������� �� �������� ������� ("`")
	
	1.2 ����������� ������� ������� ������ �� ��������� (� ���������� ������� ������������ �� �����,
	��� ��� ����������� � ��������� ����� ������ ��������� ����������)

		{"Shell, {"Folder � {"Item ������ �� �������������� �� {'Shell, {'Folder � {'Item

		"} ������ �� '}

		",[ ������ �� ',}
	
		"," ������ �� ','
	
2. �������� ������ ������ �� ���������� 

3. ������������� �����������

	3.1 ������� �������� ����� ��� ��������� ����������:

		},<����������������>  ������ �� },

	(����������� ������ � ������������ ��������� ������������� ���� ���������� �� ����� 1.2)

	3.2 � ����� ������ ������ ������ �������� ���� "\" (����� �� ��� ��������� ������� �������� ������!)

4. ������� ������� 

	4.1 �������� �������� ��������� ������� �� ���������������� ���������
	
	4.2 ��������� ������� ������� ������ �� ���� ������� �������

*/

	s = s.replace(/\\/g, '\\\\');
	
	s = s.replace(/\'/g, '`');
	
	s = s.replace(/\{\s*\"(Shell|Folder|Item)/g,	"{'$1");
	s = s.replace(/\"\s*}\s*(,|})/g,		"'}$1");

	/* ��� �������� ����:
	 *	{""��������"", ""������"", ""��������"", ""������""}
	 */
	
	// �������� ������ �������� � �����������
	s = s.replace(/\{\s*""([^\"])/g,					'{``$1');
	s = s.replace(/([^\"])\"\"\s*,\s*\"\"([^\"])/g,	'$1``,``$2');
	
	s = s.replace(/([^\"])\"\"}/g, "$1``}");
	
	s = s.replace(/\"\s*,\s*\{/g,	"',{");
	s = s.replace(/\"\s*,\s*\"/g,	"','");
	
	s = s.replace(/}\s*(,{0,1})\s+/g, "}$1");

	s = s.replace(/\{/g, '[');
	s = s.replace(/\}/g, ']');

	// ������� � ������������� �������������� ������ � als-������,
	// � �������� ����� ���������� (��� ������ Automation.als � ms_whs.als)
	s = s.replace(/\"\s*\]\s*\[\'/g, "'],['");
		
	// ����������� � �������������� ���������� �����������
	s = s.replace(/(\r\n|\r|\n)/g, "\\r\\n\\\r\n") 
		
	// ���������� ��������� � ������� ������� �� �����
	s = s.replace(/``/g, '"');
	s = s.replace(/`/g, "\\'");
	s = s.replace(/\"\"/g, '"');	

	
	// �������� �������������� ���������� ���
	try { 
		eval ("tree = " + s)
	}
	catch (e) { 
		tree = null;
		msg(e.description) 		
	}
	
	if (out) {
		out.WriteLine(s);			
		out.Close();
	}
	
	inp.Close();
	
	return tree;
}

/* ============================================================================ *\
                    ������� ���������� ������ � ���������� � XML
\* ============================================================================ */

/* processShell(alsShell)
 * 	������������ ��������������� ��� ��������� ������� �����.
 *	���������� �������������� � ���������� � XML ������.
 */
function processShell(alsShell) 
{
	var types = {}; // ����� ������ �������������� ���������� � �����
	if (alsShell[0] == 'Shell') {
		for (var i=1; i<alsShell.length; i++) {
			if (typeof(alsShell[i])=='object') {
				processFolder(alsShell[i], types);
			}
		}	
	}
	return types;
}

/* processFolder(folder, types)
 *	��������� ����������� ��������� �������.
 *	� types ����������� ���������� � �������� ��������� ���������� � �����.
 */
function processFolder(folder, types) 
{
	var typeName = folder[2];

	if (!types[typeName]) {
		types[typeName] = { 
			name	: makeValidId(folder[2]),
			alias	: makeValidId(folder[3]),
			attribs	: [], // ������ ���������
			methods	: []  // ������ �������
		}
	}

	// �������� ������ ��������� ��������� �������
	var items = folder.slice(4);

	/* ������� ��������� �������� �������:
	 *	1. ������ ��������� � ������ ������� ��� ������� ��������������� �������
	 *	2. �������� ��������� �������� � ������� "��������" ��� "������" ���������
	 *		� ��������� ��� ������� ��������������� ������������� �������
	 *	3. ��������� ������� ������������ ���������� 
	 */

	for (var i=0; i<items.length; i++) {
		if (typeof(items[i]) == 'object') {
			if (items[i][0] == 'Item') { //[1]			
				processItem(types[typeName], items[i]);
			}
			else /*if items[i][0] == 'Folder'*/ {
				var re_attr_or_meth = /^(?:��������|������|��������)$/;
				if (re_attr_or_meth.test(items[i][2])) {
					//[2] "�������������" ������ ������� �������
					var items2 = items[i].slice(4);				
					for (var j=0; j<items2.length; j++) {
						if (typeof(items2[j])=='object') {
							if (items2[j][0] == 'Folder') { // ��� ������ ��������/�����/�������� � ��� ��������
								processFolder(items2[j], types);					
							}
							else {
								processItem(types[typeName], items2[j]);
							}
						}
					}
				} 
				else { // [3] ���������� ������������ ��������� �������
					processFolder(items[i], types);
				}
			}
		}
	}	
}

/* processItem(type, item)
 *	��������� �������� ������� - ������.
 *	� type ����������� ����� ����� ��� �������� - � ����������� 
 *	�� ����������� ��������������� ������.
 */
function processItem(type, item) 
{
	if (typeof(item)=='object') {	
		if (isValidId(item[2])) { // ��� ���������� �������������?			
			if (item[4] && item[4].match(/^[�-�\w]+\(/)) { // ��� �����
				addMethod(type.methods, item);
			}
			else { // ��� ��������
				addAttrib(type.attribs, item);
			}
		}
		else {
			// � ����� ������� � ����� ����������� als-��� (ms_whs.als � ���������)
			var m = item[2].match(/^([�-�\_\w][�-�\_\w\d]+)(\(.*)/); 
			if (m) {
				item[2] = m[1];
				item[4] = item[2];
				addMethod(type.methods, item);
			}
		}
	}
}

var re_bad_ids = /(?:��������|������|����������|��������|readme)/;
var re_valid_id = /^[�-�\_\w][�-�\_\w\d]*$/;

function isValidId(id) 
{
	if (re_valid_id.test(id)) {
		return !re_bad_ids.test(id);
	}
	return false;
}

/* addMethod(methods, item)
 *	��������� ���������� � ������ � ������ ������� ����.
 */
function addMethod(methods, item) 
{	
	var data = parseDescription(item[6]);
	
	var descrStr = data['����������'];
	if (!descrStr) {
		descrStr = data['SUMMARY']?data['SUMMARY']:data['��������'];
	}

	var retStr = data['������������ ��������']?data['������������ ��������']:data['����������'];
		
	var method = {
		name 			: makeValidId(item[2]),
		alias			: makeValidId(item[3]),
		descr			: descrStr,
		params			: data['���������'],
		retTypeDescr 	: retStr?retStr.replace(/<.+?\>/,''):retStr,
		retType 		: retrieveType(retStr)
	}
	
	methods[methods.length] = method;
}

/* addAttrib(attribs, item)
 *	��������� ���������� � �������� � ������ ������� ����.
 */
function addAttrib(attribs, item) 
{
	var data = parseDescription(item[6]);
	
	var descrStr = data['����������'];
	if (!descrStr) {
		descrStr = data['SUMMARY']?data['SUMMARY']:data['��������'];
	}
	
	var attrib = {
		name 	: makeValidId(item[2]),
		alias	: makeValidId(item[3]),
		descr	: descrStr
	}
	attribs[attribs.length] = attrib;
}

// ������� (������) ��������
var re_sect = /@{0,1}(����������|���������|���������|������������ ��������|����������|������|��������|���������|[Ss]ummary|sig):*\s*/g;

/* parseDescription(source)
 *	��������� ���������� �� ���������� �������� ������.
 */
function parseDescription(source) 
{
	var data = {};

	/* ���������� ��������� ���������:
	 * 	������� �������� ����� ���������� � ������� "@", � ����� � �� ����������.
	 *	����� ����� ������� ����� ���� ��������� (":"), � ����� � �� ����.
	 *	� ��� ����� �������� ����� ������������� � ����� ������ (���) - � ��� ��� �������.
	 */
	
	// ����� ���� ����� �������� �� ����� - �� �������� ��������
	var arr = source.replace(re_sect, "<#!#>$1\r\n").split(/<#!#>/);
	
	// ��������������� ������������ ������ ������
	for (var i=0; i<arr.length; i++) {
		var lines = arr[i].split(/\r\n/);
		var id	= trim(lines[0]).toUpperCase();
		switch (id) {
			// ������ ���������� ������������ ��������, �������� ���� � ����������
			case '���������': 
				data[id] = parseParamDescr(lines.slice(1));						
				break;
		default:
			// ����� � ����� ������ ������ (������� � ��������� ����� ��������� ���)
			data[id] = lines.slice(1).join(' ').replace(/\.$/,'');
		}
	}

	return data;
}

/* parseParamDescr(lines)
 *	��������� ���������� � ���������� �� ������ ����� �������� ����������.
 *	���������� ������ ��������, �������� ���������� �� ����� ���� ���������
 *	� ��������� �������� ���������.
 */
function parseParamDescr(lines) 
{
	var m, params = [];
	for (var j=0; j<lines.length; j++) {	
		/* ������ �������� ��������� ����� ��������� ������:
		 * 	<������������> - (������������) �������� ��������� 
		 * ��� ��������� ����� �������������.
		 */		
		if (m = lines[j].match(/^<([�-�\w]+)>\s*-\s*(\(.+?\))*\s*(.+)/)) {	
			var name = m[1];	
			var type = m[2].replace(/(\)|\()/g,'').replace(/(\\|\/)/,',');
			var descr = m[3].replace(/\.$/,'');
			if (isEmpty(type)) {
				// ���� ��� �� ����� � ������� ������� � ������ ��������, ��
				// �������� ������� ��� ���� �� ������ ��������.
				type = retrieveType(descr);
			}
			params[params.length] = {
				name : name,
				type : type,
				descr: descr
			}
		}
	}
	return params;
}

/* retrieveType(str)
 *	�������� ������� ���������� � ���� �� ������ (�������) ������������� �������.
 */
function retrieveType(str)
{
	var m;
	if (!isEmpty(str)) {
		/* ��� ���� ���� ������ � �������, ��������:
		 *	(�����) 
		 */
		if (m = str.match(/^\s*\((.+?)\)/)) {
			return m[1];
		}
		/* � ������ ������ �������� ����� ���� �������� "�����" ��� "��������",
		 * "������" ��� "���������"
		 */
		if (m = str.match(/^\s*([��]���|[�c]����)/)) {
			switch (m[1].toUpperCase()) {
				case '����'	: return '�����';
				case '�����': return '������';
			}
		}
		/* ���������� � ���� ������� � ����
		 * "��� - <�������>" ��� "���: <�������>", ��������:
		 * 		��� - OLE-�����.
		 */
		if (m = str.match(/���\s*[:\-]\s*(.+?)\s*[,;\.]/)) {
			return m[1]
		}
		/* ��������, ������������ ��� �������� ������������ � 1C++
			< type="�����" >
		*/
		if (m = str.match(/<\s*type=\"(.+?)\"/)) {
			return m[1].replace(/\|/g, ',');;
		}
	}
}

// ������ 
var folder_name_patterns = [
	/������� ������ - (����)/,
	/.+?\(������ ([�-�\w]+)\)/, // ��������, "������ � ������ �� ��������� HTTP (������ V7HttpReader)"
	/������\s+([�-�\w]+)/,		// �������,  "������ ���������������"
	/.+?\((\w+)\)/				// ��������, "������ (RasDial)"
]

/* makeValidId(str)
 *	��������� �� ������ ��������� ������� ��� ����, ��������� ������������� �
 *	������� �������. ���� �� ������ ������� ������ ��������� ������� �� �������������,
 *	�� ��������� ������������� � �������� ������������� ����� �������� �������� �
 *	� ������������ �� ������.
 *	���������� - ���������� ������������� ����.
 *	TODO "��� �������������" -> "����������������" (������ ���������� "����������������")
 */
function makeValidId(str)
{
	var m;
	for (var i=0; i<folder_name_patterns.length; i++) {
		if (m = str.match(folder_name_patterns[i])) {
			return m[1];
		}
	}
	return str.replace(/\s+/g, '_').replace(/[^�-�\w\d]/g, "");
}
	
/* ============================================================================ *\
                    ������� ���������� ������ � ������� XML
\* ============================================================================ */

/* saveAsXML(types, xmlFile)
 *	��������� ����������� �� als-����� ������ � ������� xml.
 *	types - ����������� ������ � �����, ������� ��������� ���������
 *	xmlFile - ��� xml-�����, � ������� ����������� ������
 */
function saveAsXML(types, xmlFile, prefix, indent) 
{
	var xmlDoc = new ActiveXObject('MSXML2.DOMDocument');

	xmlDoc.async = false;
	xmlDoc.resolveExternals = false;

	// ��������� xml-�����
	var pi = xmlDoc.createProcessingInstruction('xml', 'version="1.0" encoding="windows-1251"');
	xmlDoc.appendChild(pi);

	// ������� � ��������� � �������� �������� ������� - ��� <Types>
	var xmlTypes = xmlDoc.createElement('Types');
	xmlDoc.appendChild(xmlTypes);

	// ��������������� ��������� ��� ����
	for (var typeName in types) {
		var type = types[typeName];
		// ���� ���� ���� �� ���� ����� ��� �������� - �� ���������
		if (type.methods.length || type.attribs.length) {
			//<type>
			var xmlType = saveTypeAsXML(type, xmlDoc, prefix);
			xmlTypes.appendChild(xmlType);
			//</type>
		}
	}

	if (indent) {
		indentXML(xmlDoc.documentElement, "\n", indent);
	}
	xmlDoc.save(xmlFile);
}

/* saveTypeAsXML(type, xmlDoc)
 *	��������� ��������� ��� type � xml-��������.
 *	���������� xml-���� �� ����������, ����������� ����������� ���.
 */
function saveTypeAsXML(type, xmlDoc, prefix) 
{
	var xmlType = xmlDoc.createElement('type');

	// �������� �������� - ��� � ������������� ���� ��� ������������
	xmlType.setAttribute('name'	,(prefix ? prefix : "") + type.name);
	xmlType.setAttribute('alias',(prefix ? prefix : "") + (isEmpty(type.alias) ? type.name : type.alias));

	// ��������� �������� (��������) ����
	if (type.attribs.length) {
		// <attribs>
		var xmlAttribs = xmlDoc.createElement('attribs');
		for (var i=0; i<type.attribs.length; i++) {
			//<attrib>
			var xmlAttrib = saveAttribAsXML(type.attribs[i], xmlDoc);
			xmlAttribs.appendChild(xmlAttrib);
			//</attrib>
		}
		xmlType.appendChild(xmlAttribs);
		//</attribs>
	}

	if (type.methods.length) {
		/* TODO �������� ����������� ������� (handlers) �� ������� (methods),
		 * ��������, ��������� ������� /^���/. XXX ������ ��� ���� ������ ��� �� 
		 * ����� ���������� ������.
		 */
		//<methods>
		var xmlMethods = xmlDoc.createElement('methods');
		for (var i=0; i<type.methods.length; i++) {
			//<method>
			var xmlMethod = saveMethodAsXML(type.methods[i], xmlDoc);
			xmlMethods.appendChild(xmlMethod);
			//</method>
		}
		xmlType.appendChild(xmlMethods);
		//</methods>
	}

	return xmlType;
}

/* saveAttribAsXML(attrib, xmlDoc)
 *	��������� ���������� � �������� (��������) ���� � xml.
 *	���������� xml-���� � ������������ �������.
 */
function saveAttribAsXML(attrib, xmlDoc) 
{
	var xmlAttrib = xmlDoc.createElement('attrib');
	xmlAttrib.setAttribute("name", attrib.name);
	if (!isEmpty(attrib.alias)) {
		xmlAttrib.setAttribute("alias", attrib.alias)
	}
	if (!isEmpty(attrib.type)) {
		xmlAttrib.setAttribute("type", attrib.type);
	}
	if (!isEmpty(attrib.descr)) {
		// ��������� �������� ��������
		var xmlDescr = xmlDoc.createTextNode(attrib.descr);
		xmlAttrib.appendChild(xmlDescr);
	}
	return xmlAttrib;
}

/* saveMethodAsXML(method, xmlDoc)
 *	��������� ���������� � ������ ���� � xml.
 *	���������� xml-���� � ������������ �������.
 */
function saveMethodAsXML(method, xmlDoc) 
{
	var xmlMethod = xmlDoc.createElement('method');
	xmlMethod.setAttribute('name', method.name);
	if (!isEmpty(method.alias)) {
		xmlMethod.setAttribute('alias', method.alias);
	}
	if (!isEmpty(method.descr)) {
		var xmlDescr = xmlDoc.createTextNode(method.descr);
		xmlMethod.appendChild(xmlDescr);
	}
	// ���� ���� ��������� - ��������� ��
	if (method.params && method.params.length) {
		//<params>
		var xmlParams = xmlDoc.createElement('params');
		for (var i=0; i<method.params.length; i++) {
			//<param>
			var xmlParam = saveMethParamAsXML(method.params[i], xmlDoc);
			xmlParams.appendChild(xmlParam);
			//</param>
		}
		xmlMethod.appendChild(xmlParams);
		//</params>
	}
	// ���������� � ������������ ��������
	if (!isEmpty(method.retTypeDescr)) {
		//<return>
		var xmlReturn = xmlDoc.createElement('return');		
		var xmlDescr = xmlDoc.createTextNode(method.retTypeDescr);
		if (!isEmpty(method.retType)) {
			xmlReturn.setAttribute('type', method.retType);
		}
		xmlReturn.appendChild(xmlDescr); // ��������� ��������
		xmlMethod.appendChild(xmlReturn);
		//</return>
	}
	return xmlMethod;
}

/* saveMethParamAsXML(param, xmlDoc)
 *	��������� ���������� � ��������� ������ � xml.
 *	���������� xml-���� � ������������ �������.
 */
function saveMethParamAsXML(param, xmlDoc) 
{
	var xmlParam = xmlDoc.createElement('param');
	xmlParam.setAttribute('name', param.name);
	if (!isEmpty(param.type)) {
		xmlParam.setAttribute('type', param.type);
	}
	if (param.descr.match(/�������������� ��������[\.,;]*|\([��]�������������\)/)) {
		// XXX ��� ������ ������� ���������� � �������� �� ���������?
		xmlParam.setAttribute("default", "");
	}
	// ��������� �������� ���������
	if (!isEmpty(param.descr)) {
		// ������ ������ ���� �� ��������
		var descr = param.descr.replace(/[��]������������� ��������[\.,;]*|\([��]�������������\)/g, '');
		var xmlDescr = xmlDoc.createTextNode(descr);
		xmlParam.appendChild(xmlDescr);
	}
	return xmlParam;
}

/* indentXML(node, curIndent, levelIndent)
 *	��������� �������������� XML-���� ���������.
 *	�����: Anton Lapounov
 *	��������: http://www.raleigh.ru/wiki/faq:save_indent
 */
function indentXML(node, curIndent, levelIndent)
{
	// ���� � ���� ��� �����������, ������ �� ������
	if (!node.hasChildNodes) return;

	// ���� ���������� �������� ��������� ����, ������ �� ������
	for (var child = node.firstChild; child != null; child = child.nextSibling) {
		if ((child.nodeType == 3)  /* PCDATA */
		||  (child.nodeType == 4)) /* CDATA */
			break;
	}
 
	if (child != null) return;
 
	// �������� ��������� ���� � ��������� ����� ��������� ������ � ���������� �� ����������
	var newIndent = curIndent + levelIndent;
	var textNode = node.ownerDocument.createNode(3, "", "");
	textNode.text = newIndent;

	for (child = node.firstChild; child != null; child = child.nextSibling) {
		node.insertBefore(textNode.cloneNode(false), child);
		indentXML(child, newIndent, levelIndent);
	}
 
	textNode.text = curIndent;
	node.appendChild(textNode);
}
/* ============================================================================ *\
                    ������� ���������� ������ � INTS-������
\* ============================================================================ */

function saveAsInts(types, intsFolder, prefix) 
{	
	var crObjFile = fso.OpenTextFile(fso.BuildPath(intsFolder, "�������������.ints"), 8, true);
	for (var typeName in types) {
		var type = types[typeName];
		if (type.attribs.length||type.methods.length) {
			var validTypeName	= (prefix ? prefix+"." : "") + makeValidId(typeName);			
			var intsFileName	= fso.BuildPath(intsFolder, validTypeName) + ".ints";
			crObjFile.WriteLine("0000 " + validTypeName);
			msg(intsFileName + " ...");			
			var intsFile = fso.OpenTextFile(intsFileName, 2, true);
			saveTypeAsInts(types[typeName], intsFile);
			intsFile.Close();
		}
	}
	crObjFile.Close();
}

function saveTypeAsInts(type, intsFile) 
{
	// ��������...
	var attribs = type.attribs;
	if (attribs.length) {
		for (var i=0; i<attribs.length; i++) {
			intsFile.WriteLine("0000 " + attribs[i].name);
		}
	}
	// ������...
	var methods = type.methods;	
	if (methods.length) {
		for (var i=0; i<methods.length; i++) {
			var f = (methods[i].params && methods[i].params.length)?"f":"";
			intsFile.WriteLine("0000 " + methods[i].name + "(" + f + ")");
		}
	}
}

/* ============================================================================ *\
                          ��������������� �������
\* ============================================================================ */

function trim(str) 
{
	if (!str) return "";
	return str.replace(/^\s+/,'').replace(/\s+$/,'');
}

function isEmpty(s) 
{
	return (!s || /^\s*$/.test(s));
}

function msg(str, marker)
{
	if (!Silent) {
		if (OpenConf) {
			Message(str, marker?marker:mNone);
		}
		else {
			WScript.Echo(str)
		}
	}
}

function exit(errCode)
{
	WScript.Quit(errCode);
}

function inConfigurator(_)
{
	try {
		var version = Configurator.Version;
		return true;
	} catch(e) {
		return false;
	}
}

/* ============================================================================ *\
                               �������� ���������	
\* ============================================================================ */

OpenConf = inConfigurator();

if (OpenConf) 
	Init() 
else {
	Run();
	exit(0);
}

