$NAME ���������

/*===========================================================================
Copyright (c) 2004 Alexander Kuntashov
=============================================================================
������:  navigator.js ("���������")
������:  1.2
�����:   ��������� ��������
E-mail:  kuntashov at yandex dot ru
ICQ UIN: 338758861
��������: 
	������ "���������" ��� OpenConf ������������ ��� �������� ������� � ��������
���������� ������������ ��� �������� �������� �������, ���������� � ������ ������
������ (����������) ��� ������������� (������ ��� ����� �������� �������) �����. 
===========================================================================*/

/* ��������� �� ������ */
function error(str)
{
	with (new ActiveXObject("WScript.Shell"))
		Popup(str, 0, SelfScript.Name, 0|48);
}

/* ���������� ������ TextDoc, ���� ������� �������� ���������� ��������
	��������� ��������, � ��������� ������ - null */
function getTextDoc(_)
{
	var win, doc;

	if (!(win = Windows.ActiveWnd))
		return null;

	if ((doc = win.Document) == docWorkBook)
		doc = doc.Page(1);

	return (doc == docText)?doc:null;
}

/* ������������� ������ ��� �������� ��� ����, ��������� ���
	���� openDialog - true, �� ��� ���������� �������� ������������ �
	������� ���������/������� ��� �������� ������������ �������� "������",
	� ��������� ������ - �������� "������" */
function openPath(openDialog)
{
	var doc, cl, re, path, m, kind, dot;
	
	if(!(doc = getTextDoc()))
		return;

	cl = doc.Range(doc.SelStartLine);   

	re = /[.#\w�-�\\:\/]/i;

	path = "";

	if (doc.SelStartCol != doc.SelEndCol)
		path = cl.substring(doc.SelStartCol, doc.SelEndCol);
	else {
		for (var pos=doc.SelStartCol; 
				(pos>0) && cl.charAt(pos).match(re);)          
			path = cl.charAt(pos--) + path;       
	
		for (pos=doc.SelStartCol+1; 
				(pos < cl.length) && cl.charAt(pos).match(re);)        
			path += cl.charAt(pos++);   
	}

	try {
		if (path.match(/\\|\//)) {
			openFileSystemPath(path, openDialog);
		}
		else {
			if (!openMetaDataPath(path, openDialog)) {
				error('�� ���� ������� ������: ' + path);
			}
	   }
	} 
	catch (e) {
		error(e.description);
	}  
}

/* ��������� ������� ����/����� */
function openFileSystemPath(p, openDialog)
{
	var path;
		   
	path = p.match(/^\w:/)?p:CommonScripts.ResolvePath(IBDir, p);            

	with(new ActiveXObject("Scripting.FileSystemObject")) {               
		if (FileExists(path)) {
			openFile(path, openDialog);                            
		} 
		else if (FolderExists(path)) {
			openFolder(path);                                    
		} 
		else
			error('���� "' + path + '" �� ����������!');
	}  
}

/* ��������� ������� ���� */
function openFile(path, openDialog)
{
	var ext;
	   
	with (new ActiveXObject("Scripting.FileSystemObject")) 
		ext  = GetExtensionName(path).toLowerCase();       
	 
	switch (ext) {
		case 'txt': case 'ert': case 'mxl':
		case '1s':  case '1c':  case '1�':
			/* ��������� ���� � ������������� */
			openFileInside(path, openDialog);
		break;
		/* �� ��������� �������� ������� ���� 
		   ��������������� � ������� ����� ���� ���������� */
		default:
			openFileOutside(path, ext);
	}
}

/* ��������� ������� ���� � ������������� */
function openFileInside(path, openDialog)
{
	with (Documents.Open(path)) 
		if ((Type == docWorkBook)&&(!openDialog))             
			ActivePage = 1;                                 
}

/* ��������� ������� ���� � ������� ���������, ���������������
	� ������� � ��������������� ����� ������ */
function openFileOutside(path, ext)
{
	var key, cmdTmpl, cmd;
	with (new ActiveXObject("WScript.Shell")) {
		try {            
			key     = RegRead("HKCR\\." + ext + "\\");                    
			cmdTmpl = RegRead("HKCR\\" + key + "\\shell\\open\\command\\");            

			cmd = cmdTmpl.replace(/%1/, path);
			if (cmd == cmdTmpl)
				cmd = cmdTmpl + ' ' + path;                        

			Run(cmd);
		} 
		catch(e) {
			//error(e.description);
			error("��� �������� ����� ����� ���� " + 
				"�� ������������ ������� ���������!");
		}        
	}
}

/* ��������� ����� � ���������� */             
function openFolder(path)
{
	with (new ActiveXObject("WScript.Shell"))
		Run("%SYSTEMROOT%\\explorer.exe " + path);
}

/* ��������� ������������� ������� ������������ 
   � ����� type � ����� name */
function findMDO(type, name)
{
	var obj, mdObjs = MetaData.TaskDef.Childs;    
	for (var i = 0; i < mdObjs(type).Count; i++) {
		if (mdObjs(type)(i).Name.toUpperCase() == name.toUpperCase()) {
			return mdObjs(type)(i);
		}
	}    
	return  null;
}

/* ���������� ������������, ����� ����� 
   ����������� ������� ������� */
function askRefForm(name)
{    
	var ref, forms = new Array();
	
	if (!(ref = findMDO("����������", name))) {         
		return null;
	}
		
	for (i = 0; i < ref.Childs("�����������").Count; i++) {
		forms[i] = "�����������." + ref.Childs("�����������")(i).Name;
	}

	with (new ActiveXObject("Svcsvc.Service"))
		return SelectValue("�����\r\n" + "����� ������\r\n" 
				+ ((forms.length>0)?forms.join("\r\n"):""), 
				'����� - ����������.' + name);
}

/* ���������� ������������, ����� ����� ������ 
   ������� ���������� ������� ������� */
function askJournalForm(name)
{        
	var jrn, forms = new Array();
	
	if (!(jrn = findMDO("������", name))) {         
		return null;
	}
	   
	for (i = 0; i < jrn.Childs("�����������").Count; i++) {
		forms[i] = jrn.Childs("�����������")(i).Name;
	}

	if (forms.length == 0) {
		return null;
	}
	if (forms.length == 1) { 
		// ����� ������ ����, ���������� �� ����
		return forms[0];
	}
	
	with (new ActiveXObject("Svcsvc.Service"))
		return SelectValue(forms.join("\r\n"), '����� ������ - ������' + name);
}

/* ���������� ������������, ��������� ��� 
   ��������� ������ ����� ��� ������ ���������� */
function askDocModule(name) 
{
	var a = new Array("������ ���������", "�����.������"); 

	with (new ActiveXObject("Svcsvc.Service"))
		return SelectValue(a.join("\r\n"), "������ - ��������." + name);
}

/* ��������� �����/������ ��� ������� ���������� */
function openMetaDataPath(p, openDialog)
{
	var m, dot, path, tmp;

	/* ��� ��������� ������������� �������������� ���� */    
	dot = new Function('s', 'return "." + s;');

	switch ((m = p.split('.')).length) {
		case 0: 
			return false;
		case 1:
			if (m[0].toUpperCase() == '��������')
				break;            
			if (findMDO("����������", m[0])) {
				m = new Array("����������", m[0]);
			}
			else if (findMDO("��������", m[0])) {
				m = new Array("��������", m[0]);
			}
			else if (findMDO("������", m[0])) {
				m = new Array("������", m[0]);
			}
			break; 
		default:
			// do nothing :-)
	}
		
	switch (m[0].toUpperCase()) {                
		/*---------------------------------------------------------- 
			(�������|����������).��� / ����������.���.��������������  
		----------------------------------------------------------*/
		case '����������': case '�������':                                                       
			if (!(tmp = (m.length == 2)?askRefForm(m[1]):("�����������." + m[2]))) {
				// ������������ ������� �������� ������/������� �����
				return true; 
			}
			path = '����������' + dot(m[1]) + dot(tmp);             
		break;
		/*---------------------------------------------------------- 
			��������.������������
		----------------------------------------------------------*/
		case "��������": 
		//	tmp = "�����.������";
		//	if (!openDialog) {
		//		if (!(tmp = askDocModule(m[1]))) {               
		//			// ������������ ������� �������� ������/������� �����
		//			return true;
		//		}               
		//	}
			tmp = ((getTextDoc().Name.toUpperCase().search('������ ���������') !== -1)?(openDialog?'�����.������':'�����.������'):'������ ���������')
			path = m[0] + dot(m[1]) + dot(tmp);
		break;        
		/*---------------------------------------------------------- 
			(�����|���������).�������������
		----------------------------------------------------------*/
		case "�����": case "���������":
			path = m[0] + dot(m[1]) + dot("�����");
		break;
		/*---------------------------------------------------------- 
			������.���������.��������������
		----------------------------------------------------------*/
		case "������":
			if (m[1].toUpperCase() == "�����������") {
				// ���� ���� �� ������� �� ����� ���� �����, 
				// �� � ������� �� �����, ����� �� ������� ������
				// � �����������
				return false; 
			}
			if (!(tmp = (m.length == 2)?askJournalForm(m[1]):m[2])) {
				// ������������ ������� �������� ������/������� �����
				return true;
			}                                     
			path = '������' + dot(m[1]) + dot('�����') + dot(tmp);
		break;
		/*---------------------------------------------------------- 
			��������������.������������������(.��������������)*
		----------------------------------------------------------*/
		case "��������������":
			path = '��������������' + dot(m[1]) + dot('�����') + 
				dot((m.length == 3)?m[2]:'�����������') + dot('�����');
		break;
		/*---------------------------------------------------------- 
			����.��������������
		----------------------------------------------------------*/
		case "����":
			path = '����������' + dot('�����');
		break;
		/*---------------------------------------------------------- 
			����������.�������������.�������������
		----------------------------------------------------------*/
		case "����������":
		  path = '����������' + dot('�����������') + dot(m[2]) + dot('�����');
		break;  
		/*---------------------------------------------------------- 
			��������
		----------------------------------------------------------*/
		case "��������":           
			path = '��������' + dot('�����');
		break;
		/*---------------------------------------------------------- 
			��������������.����������������������
		----------------------------------------------------------*/
		case "��������������":
			path = '��������.�����������' + dot(m[1]) + dot('�����');
		break;
		/*---------------------------------------------------------- 
			��������������.����������������������
		----------------------------------------------------------*/
		case "��������������":
			path = '��������' + dot('�����������') + dot(m[1]) + dot('�����');            
		break;
		/*---------------------------------------------------------- 
			�������.���������.�����������, 
			�������.����������.��������������.�������������������
			�������.����
		----------------------------------------------------------*/
		case "�������":
			/* ���� �� ���� (�� �����), 
				��� ����� ������� ����� ������� � ������������� */
			error("��������, ���� �� �����������. ����: " + m.join('.'));
		//break;        
		/*---------------------------------------------------------- 
			��� ��������� ����������� ����������
		----------------------------------------------------------*/
		default:        
			return;
	}

	if (m[0].toUpperCase() !== "��������") {
		path += dot(openDialog?'������':'������');
	}
 
	Documents(path).Open();           
	return true;
}

/* ��� ������� ��������� ������ ������� (������ ����� �� 
	���� ������� ��� ������ ���������� ���������, � ������,
	����� ��������������� ������ �������� ����������, ���������� 
	������������ ������� �� ������ � ������� ����� ������ ��������� 
	��� ����� ������� ������ ��� ����� */
function openAnotherModuleOrFormForCurrentObject(openDialog)
{
	var doc, m;       

	if (!(doc = Windows.ActiveWnd))
		return null;

	doc = doc.Document;
	
	m = doc.Name.split('.');

	switch (m[0].toUpperCase()) {
		case '��������':
		case '����������':
		case '������':
			openMetaDataPath(m[0] + '.' + m[1], openDialog);    
			break;
		default:
			; // do nothing
	}
}

//===========================================================================
//                          user interface :-)
//===========================================================================

/* ������� ��� ��������� ��������� */

// Ctrl + Shift + M[odule]
function ����������������������������������()
{
	openPath();
}

// Ctrl + Shift + F[orm]
function ���������������������������������()
{
	openPath(true);
}

// Alt + Shift + M[odule]
function �������������������������������������()
{
	openAnotherModuleOrFormForCurrentObject();
}

// Alt + Shift + F[orm]
function ������������������������������������()
{
	openAnotherModuleOrFormForCurrentObject(true);
}

/* ��������� ��������������� ������ � ����������� ���� (��. �������� � 
	������ "�������" - ���� ������ ���) ��� ������ �������� */
function TelepatGetMenu(_)
{
	var menu = "������� ������...| |__NavigatorJS�������������\r\n" + 
			   "������� �����... | |__NavigatorJS�������������\r\n";        
	return menu;
}

/* ���������� ������ ������ �� ������� ������ ���� */
function TelepatOnCustomMenu(cmd)
{
	if (cmd == '__NavigatorJS�������������')
		openPath(true);
	if (cmd == '__NavigatorJS�������������')
		openPath(false);
}

//===========================================================================

/* ��� ��������� ����������� ������ JScript'� � 
	������������ ������������ ������� ��������, �����������
	� ������������ ���� ������� run-time -- a13x */ 
function addTelepatEventHandler(handler)
{
	eval('function Telepat::' + handler + '{ return Telepat' + handler + '}');
}

/*
 * ��������� ������������� �������
 */
function Init(_) // ��������� ��������, ����� ��������� �� �������� � �������
{
	try {
		var t, ocs = new ActiveXObject("OpenConf.CommonServices");
		ocs.SetConfig(Configurator);
		SelfScript.AddNamedItem("CommonScripts", ocs, false);

		/* ��������� ������ ������� � ������������ ���� ������� */
		
		ocs.ClearError();
		ocs.SetQuietMode(true);		
		ocs.AddPluginToScript(SelfScript, "�������", "Telepat", t);		
   		ocs.SetQuietMode(false);				
		
		if (ocs.GetLastError()) {	
			var warnUser = true, OCReg = ocs.Registry;
			if (OCReg !== null) {
				var srk = OCReg.ScriptRootKey(SelfScript.Name);				
				if (OCReg.Param(srk, "DoNotWarnOnStartUp") == "true") {
					warnUser = false;
				}
				else {
					// ��� ��������� ������� ������������� ��������� �� ������ ���������
					OCReg.Param(srk, "DoNotWarnOnStartUp") = "true";
				}
			}	
			if (warnUser) {
				Message(ocs.GetLastError(), mRedErr);			
				Message("��������, � ��� �� ���������� ������ ������� 2");
				Message("����� �������� ������� ��������� �� ������������ ���� ����� ����������!");
			}
		}
		else {
			/* ���������� ������ ������ ��� ����������� ������� �������� */
			
			//����� ���� ��������
			addTelepatEventHandler("GetMenu()");

			//����� ������ ���� ��������
			addTelepatEventHandler("OnCustomMenu(cmd)");
		}	

	}
	catch (e) {
		Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("������ " + SelfScript.Name + " �� ��������", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}	
}

//===========================================================================
		 
/* �������... */
Init();                                                     
