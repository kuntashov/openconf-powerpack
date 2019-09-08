function CreateObjectOrDie(progid, id) 
{
	try {
		var obj = new ActiveXObject(progid);
		if (id) {
			SelfScript.AddNamedItem(id, obj, false);		
		}
		return obj;
	}
	catch (e) {
		Message("�� ���� ������� ������ " + progid, mRedErr);
		Message(e.description, mRedErr);
		Message("������ " + SelfScript.Name + " �� ��������", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}		
}

function MakeList() {

	var htm = '<html><head><title>������ �������������� ������������ ����������</title>'
			+ '<style type="text/css">'
			+ '.lightRow{ font-size:10pt; font-family:"Courier New"; background-color:rgb(234,234,234); } '
			+ '.darkRow	{ font-size:10pt; font-family:"Courier New"; background-color:rgb(204,204,255); } '
			+ '.tHeader { font-size:10pt; font-family:"Courier New"; font-weight:bold; } '
			+ '</style>'
			+ '</head><body><center>'
			+ '<table width="70%">'
			+ '<tr class="tHeader"><td>������</td><td>������</td><td>������</td></tr>';

	ocKM.Load();

	var scans = (new VBArray(ocKM.Codes())).toArray();	

	for (var i=0; i<scans.length; i++) {

		var hki		= ocKM.HotKeyInfo(scans[i]);
		var mnemo	= smKM.toShortCut(scans[i], "+", true);		

		htm +=	'<tr class="' + ((i%2)?'lightRow':'darkRow') + '">'
			+	'<td>' + mnemo		+ '</td>'
			+	'<td>' + hki.Modul	+ '</td>'
			+	'<td>' + hki.Macros	+ '</td>'
			+	'</tr>';
	}

	htm += "</table></center></body></html>";

	// ������� ������ � �������������

	var fso = new ActiveXObject("Scripting.FileSystemObject");
	var fname = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName());
	var f = fso.CreateTextFile(fname, true);
	f.Write(htm);
	f.Close();

	var wnd = OpenOleForm("Shell.Explorer", "HotKeys List");
	wnd.Navigate2(fname);	
	
}

CreateObjectOrDie("OpenConf.RegistryKeyMap", "ocKM");
CreateObjectOrDie("ScriptManager.KeyMap", "smKM");
