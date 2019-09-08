$NAME ������������� �����

/*===========================================================================
������:  MultiString.js
������:  $Revision: 1.1 $
�����:   ��������� ��������
E-mail:  kuntashov at [yandex.ru|gmail.com]
ICQ UIN: 338758861
��������: 
	������� ��� ������� �� ������/����������� � ����� ������������� 
	��������� �������� �� ������� 1� � �����������/��������� �������� ��������
	����� ������ (|) � ������� ��������� ������� ������� (") �� ������ ("") 
	� �������.
===========================================================================*/

function ����������������������() {

	var doc	= CommonScripts.GetTextDocIfOpened();	
	var str = CommonScripts.GetFromClipboard();
	
	if (doc && str) {
		var m, ind = "";
		/* ��� ��������� ����� �������������� ������ ���������� 
			������ ������, � ������� ��������� ������ */
		if (m=doc.Range(doc.SelStartLine).match(/^(\s+)/)) {
			ind = m[1];
		}
		doc.Range(
			doc.SelStartLine, doc.SelStartCol, 
			doc.SelEndLine, doc.SelEndCol) = str.replace(/\n/g, "\n"+ind+"\|").replace(/\"/g,'""');
	}
}

function ����������������������() {

	var doc	= CommonScripts.GetTextDocIfOpened();	

	if (doc) {
		var str = doc.Range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol);
		if (str) {
			CommonScripts.CopyToClipboard(str.replace(/^\s*\|/gm, "").replace(/\"\"/g, '"'));
		}
	}	
}

/*
 * ��������� ������������� �������
 */
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

Init();

