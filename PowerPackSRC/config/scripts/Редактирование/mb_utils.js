$NAME Metabuilder-�������
//$Id: mb_utils.js,v 1.7 2005/09/07 10:23:43 alest Exp $
//
//����� ������� ������� aka alest (j_alesant@mail.ru)
//��������: 
//  ����� ������� aka artbear
//  a13x
//
//���������������� �� �������� GPL

//������� � ������ ����� ���� � �������
//���������� ������ � ������� ���������� �����������
function cleanEOL(doc) {
	var re = /\s+$/;
	var lines = doc.Text.split("\r\n");
	var corrections = 0;
	var bChangeSel = 0;
	for (var i = 0; i < lines.length; i++) {
		var oldLength = lines[i].length;
		var selStartCol = doc.SelStartCol;
		var selEndCol = doc.SelEndCol;
		lines[i] = lines[i].replace(re, "");
		if (lines[i].length !=  oldLength) {
			doc.range(i, 0, i, oldLength) = lines[i];
			if ((i == doc.SelStartLine) && (selStartCol > lines[i].length)) {
				selStartCol = lines[i].length;
				bChangeSel = 1;
	}
			if ((i == doc.SelEndLine) && (selEndCol > lines[i].length)) {
				selEndCol = lines[i].length;
				bChangeSel = 1;
			}
			//Message("������ #" + (i+1));
			corrections++;
		}
	}
	if (bChangeSel) {
		doc.MoveCaret(doc.SelStartLine, selStartCol, doc.SelEndLine, selEndCol);
	}
	if (corrections)
		Message("�������� �������� � ��������� � ����� �����: " + corrections);
}

//������ �� Ctrl-S
//��� ������� ����� ������� ������� � ������ ����� ���� � �������
function save() {
	var a;
	if (a = Windows.ActiveWnd) {
		if (a = a.Document) {
			if (a == docWorkBook && a.ActivePage == 1) {
				a = a.Page(1);
			}
			if (a == docText) {
				cleanEOL(a);
			}
		}
	    SendCommand(cmdSave);
	}
}

//������������� ����� ����� (������ �������� ����� � �������� ����)
//��� �������, ����� ���� �������� ������ �����������
function reloadFile(doc) {
	if (CommonScripts.FSO.FileExists(doc.Path)) {
		var stream = CommonScripts.FSO.OpenTextFile(doc.Path, 1, false);
		var text = stream.ReadAll();
		stream.Close();
		doc.Text = text;
		doc.Save();
	}
}

//��������� ������������ ������ ��� ��������� ����
function reloadActiveWnd() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		reloadFile(doc);
	}
}

//��������� ������������ ��� ���� �������� ��������� ����
//����� cvs update �������� �������
function reloadAllFiles() {
    var wnd = Windows.FirstWnd;
    while (wnd) {
        var doc = wnd.Document;
		if (doc == docText) {
			reloadFile(doc);
		}
        wnd = Windows.NextWnd(wnd);
	}
}

function closeActiveWnd() {
	if (Windows.ActiveWnd) {
		Windows.ActiveWnd.Close();
	}
}

//��������� ����������������.txt  � �� � ������� glob2md.bat
function glob2md() {
	var wsh = new ActiveXObject("WScript.Shell");
	wsh.CurrentDirectory = IBDir;
	var CmdLine = "glob2md.bat";
	var oExec = wsh.Exec(CmdLine);
	var outputLines = oExec.StdOut.ReadAll().split("\r\n");
	Message("Glob2md " + outputLines[outputLines.length-2]);
}

function start1C() {
	glob2md();
	Run1CApp(rmEnterprise);
}

//���������� � ����� ���� ��������� � �������� ���� �����
function copyFilePath() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		CommonScripts.CopyToClipboard(doc.Path);
	}
}

//�������� ������� ��� ������� ������������ ����� ��������� �� �������
//���� ��������������� �������� � ������� �� �����, �� ������������
//������� ������ �������� (����� ����������� � �������)
function getEditorCmd(defaultEditor) {
	var editor = defaultEditor;
	var ocReg = CommonScripts.Registry;
	if (!ocReg) {
		Message(CommonScripts.GetLastError(), mRedErr);
		return defaultEditor;
	}
	try {
		var rk = ocReg.ScriptRootKey(SelfScript.Name);
		editor = ocReg.Param(rk, "Editor");
		if (!editor) {
			editor = CommonScripts.SelectFileForRead("", "����������� ����� (*.exe)|*.exe");		
			if (editor) {
				ocReg.Param(rk, "Editor") = editor;
			}
		}
	}
	catch (e) {
	}
	return editor ? editor : defaultEditor;
}

//������� ���� ��������� ���� � ������ ���������
function openFileWithEditor() {
	var editor = getEditorCmd("notepad.exe");
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		CommonScripts.RunCommand(editor, doc.Path, 0);
	}
}

//��������� ���� � ���������� ����������������(LoadFromFile) � ����������������.txt (�� ���� ������������ � �� ����. ������)
function openIncludeFile() {
	var a = Windows.ActiveWnd;
	if (!a) {
		SendCommand(cmdOpenConfigWnd);
	}
	else if (a && a.Caption.indexOf("������������") != -1) {
	    if (!Documents.Open(IBDir + "Modules\\ModuleText\\����������������.txt")) {
	    	Documents("����������������").Open();
	    }
	}
	else {
		var doc = CommonScripts.GetTextDoc(0);
		if (doc) {
			if (doc.Kind == "ModuleText") {
			    var d = Documents.Open(IBDir + "Modules\\ModuleText\\����������������.txt");
				d.MoveCaret(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol);
			}
			else {
				var firstLine = doc.range(0, "\r\n");
				if ( (firstLine.toUpperCase().indexOf("#����������������".toUpperCase())  != -1)
				|| (firstLine.toUpperCase().indexOf("#LoadFromFile".toUpperCase()) != -1) ) {
					firstLine = firstLine.replace(/^\s*(\S.*\S)\s*$/, "$1");
					var path = firstLine.substring(18);
					var d = Documents.Open(IBDir + path);
				    if (!d) {
						d = Windows.ActiveWnd.Document;
						if (d == docWorkBook) {
							if (d.ActivePage != 1)
								d.ActivePage = 1;
				    	}
				    }
				}
				else {
					var d = Windows.ActiveWnd.Document;
					if (d == docWorkBook) {
						if (d.ActivePage != 1)
							d.ActivePage = 1;
			    	}
				}
			}
		}
	}
}

//�������������� ����� 
//� ������� �� �������� ������ �����������, 
//���� ���� ������ �� �������� ��� ������� ����� �� ����� ������
function commentSelection() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		var sCommentSymbol = "//";
		//���� ������ �� �������� ��� ������� ����� �� ����� ������
		if (doc.SelStartLine == doc.SelEndLine) {
			var str = doc.Range(doc.SelStartLine);
			var pos = 0;	//������� �������� �����������
			var selection = {startCol: doc.SelStartCol, endCol: doc.SelEndCol};
			if (str) {
				var matches = str.match(/^(\s*)(\S.+)/);
	            if (matches.length) {
					doc.Range(doc.SelStartLine) = matches[1] + sCommentSymbol + matches[2];
					pos = matches[1].length;
        		}
        	}
			else {//���� �����, ����������� ������ � ������
				doc.Range(doc.SelStartLine) = sCommentSymbol + doc.Range(doc.SelStartLine);
			}
			if (pos <= doc.SelStartCol) {
				selection.startCol += sCommentSymbol.length;
			}
			if (pos <= doc.SelEndCol) {
				selection.endCol += sCommentSymbol.length;
			}
			doc.MoveCaret(doc.SelStartLine, selection.startCol, doc.SelStartLine, selection.endCol);
		}
       	else {
			doc.CommentSel();
		}
	}
}
         
//����������������� ����� 
//� ������� �� �������� ������ �����������, 
//���� ���� ������ �� �������� ��� ������� ����� �� ����� ������
function unCommentSelection() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		var sCommentSymbol = "//";
		//���� ������ �� �������� ��� ������� ����� �� ����� ������
		if (doc.SelStartLine == doc.SelEndLine) {
			var selection = {startCol: doc.SelStartCol, endCol: doc.SelEndCol};
			var str = doc.Range(doc.SelStartLine);
			var pos = str.indexOf(sCommentSymbol); 	//������� �������� �����������
			if (pos != -1) {
				doc.Range(doc.SelStartLine) = str.substring(0, pos) + str.substring(pos+sCommentSymbol.length);

				if (pos < selection.startCol) {
					selection.startCol -= (sCommentSymbol.length - Math.max(pos + sCommentSymbol.length - selection.startCol, 0));
				}
				if (pos < selection.endCol) {
					selection.endCol -= (sCommentSymbol.length - Math.max(pos + sCommentSymbol.length - selection.endCol, 0));
				}
				doc.MoveCaret(doc.SelStartLine, selection.startCol, doc.SelStartLine, selection.endCol);
			}
	    }
		else {
			doc.UnCommentSel();
		}
	}
}

var PositionInModule;
//�������� ��������� ����� ������, ��������������� ��������
function getNextLineMatch(lines, re, curPosition) {
	for (var i = curPosition; i < lines.length; i++) {
		var str = lines[i].toUpperCase();
		if (str.search(re) != -1) {
			return i;
		}
	}
	return -1;
}

//�������� ����������� ����� ������, ��������������� ��������
function getPrevLineMatch(lines, re, curPosition) {
	for (var i = curPosition; i >= 0; i--) {
		var str = lines[i].toUpperCase();
		if (str.search(re) != -1) {
			return i;
		}
	}
	return -1;
}

function getFunctionStart(text, curPosition) {
	return getPrevLineMatch(text.split("\r\n"), /^\s*�������|^\s*���������/, curPosition);
}

function getFunctionEnd(text, curPosition) {
	return getNextLineMatch(text.split("\r\n"), /^\s*������������|^\s*��������������/, curPosition);
}

function gotoFunctionStart() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		PositionInModule =  CommonScripts.GetDocumentPosition(doc);
		var i = getFunctionStart(doc.Text, doc.SelStartLine);
		if (i >= 0) {
			doc.MoveCaret(i, 0);
		}
	}
}

function gotoFunctionEnd() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		PositionInModule =  CommonScripts.GetDocumentPosition(doc);
		var i = getFunctionEnd(doc.Text, doc.SelStartLine);
		if (i >= 0) {
			doc.MoveCaret(i, 0);
		}
	}
}

function gotoSavedPosition() {
	PositionInModule.Restore();
}

function selectCurrentFunction() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	if (doc) {
		doc.MoveCaret(getFunctionStart(doc.text, doc.selStartLine), 0, getFunctionEnd(doc.Text, doc.selStartLine)+1, 0);
	}
}

function init(_) // ��������� ��������, ����� ��������� �� �������� � �������
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

init(0);
