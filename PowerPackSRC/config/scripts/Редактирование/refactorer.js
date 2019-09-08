$NAME 1�-����������
//$Id: refactorer.js,v 1.2 2005/09/27 20:45:07 alest Exp $
//
//����� ������� ������� aka alest (j_alesant@mail.ru)
//��������: 
//
//���������������� �� �������� GPL

var mTextEditor;
var mTextAnalyzer;

//����� ��� �������������� ������ � ���������� ������ ������ ������
function TextEditor(_) {
	var doc;
	
	this.getDoc = function () {
		return doc;
	};
	
	this.setDoc = function (lDoc) {
		doc = lDoc;
	};

	//�������� ����� ��������� � ������� ���� ��� ������ �����
	this.getDocText = function (asArray) {
		if (asArray)
			return this.text2LinesArray(doc.Text);
		else
			return doc.Text;
	};
	
	//�������� ����� ��������� ����� � ������� ���� ��� ������ �����
	this.getText = function (startLine, startCol, endLine, endCol, asArray) {
		if (asArray) {
			var lines = this.getDocText(true).slice(startLine, endLine);
			return lines;
		}
		else {
			return doc.range(startLine, startCol, endLine, endCol);
		}
	};
	
	this.getSelectedText = function () {
		return doc.range(doc.SelStartLine, doc.SelStartCol, doc.SelEndLine, doc.SelEndCol);
	};
	
	this.text2LinesArray = function (text) {
		return text.split("\r\n");
	};
	
	this.getSelectionHeight = function () {
		//Message(doc.SelStartLine + "-" + doc.SelEndLine + ";" + doc.SelStartCol + "-" + doc.SelEndCol);
		if (doc.SelStartLine == doc.SelStartCol) return 1;
		var height = 0;
		if (doc.SelEndCol == 0) height--;
		height += doc.SelEndLine - doc.SelStartLine + 1;
		//Message("height=" + height);
		return height;
	};
	
/*	this.getTextHeight = function (text) {
		var lines = text.split("\r\n");
		return lines.length;
	};
	
	//check if text has \r\n characters
	this.textHasCR = function (text) {
		if (text.search(/\r\n/) != -1) 
			return true;
		else
			return false;
	};
*/	
	//���������� ����� � ��������� �������
	this.setText = function (text, startLine, startCol, endLine, endCol) {
		doc.range(startLine, startCol, endLine, endCol) = text;
	};
	
	//������� �� ������ ������������������ �����
	this.trimComments = function (text) {
		var commentStartPos = text.search(/\/\//);
		if (commentStartPos == -1) {
			return text;
		}
		else {
			return text.substr(0, commentStartPos);
		}
	};
	
	this.insertLine = function(atLine) {
		if (atLine > doc.LineCount) {
			atLine = doc.LineCount;
		}
		this.setText(this.getText(atLine, 0, atLine, null, false) + "\r\n"
		, atLine, 0, atLine, null);
		//��������� ���������
		doc.MoveCaret(doc.SelStartLine+1, doc.SelStartCol, doc.SelEndLine+1, doc.SelEndCol);
	};

	this.insertLines = function (lines, atLine) {
		if (atLine > doc.LineCount) {
			atLine = doc.LineCount;
		}
		for (var i = 0; i < lines.length; i++) {
			var curLine = atLine + i;
			this.insertLine(curLine, lines[i]);
			curLine++;	//���������� ����� �����������
			this.setText(lines[i], curLine, 0, curLine, 0);
		}
	};
	
	this.insertText = function (text, atLine) {
	};
	
/*	//�������� ���� �� ��������
	this.deleteLine = function (atLine) {
		if (atLine > doc.LineCount) {
			atLine = doc.LineCount;
		}
		this.setText("", atLine, 0, atLine, null);
		//��������� ���������
		doc.MoveCaret(doc.SelStartLine-1, doc.SelStartCol, doc.SelEndLine-1, doc.SelEndCol);
	};
	
	this.deleteLines = function (atLine, number) {
		var i = atLine;
		var lastIndex = atLine + number - 1;
		for (var i = atLine; i <= lastIndex; i++) {
			if (i > doc.LineCount) break;
			Message("Count=" + doc.LineCount);
			this.deleteLine(i);
		}
	};*/
}

//����� ��� ������� ������
function TextAnalyzer(textEditor) {
	//�������� ������������ ����������� ������ � ��������� ������
	this.getMethodCallAtLine = function (atLine) {
		var text = textEditor.getText(atLine, 0, atLine, null, false);
		text = textEditor.trimComments(text);
		var re = /([�-�a-z]+)\s*\(.*\)/gi;
//		testRe(text, re);
		var matches = re.exec(text);
		if (matches) {
			return matches[1];
		}
		else {
			return "";
		}
	};
	
	//�������� ��������� ����� ������, ��������������� ��������
	function getNextLineMatch(lines, re, curPosition) {
		for (var i = curPosition; i < lines.length; i++) {
			if (lines[i].search(re) != -1) {
				return i;
			}
		}
		return -1;
	};

	//�������� ����������� ����� ������, ��������������� ��������
	function getPrevLineMatch(lines, re, curPosition) {
		for (var i = curPosition; i >= 0; i--) {
			if (lines[i].search(re) != -1) {
				return i;
			}
		}
		return -1;
	};

	this.getMethodStartLine = function (curPosition, bWithComments) {
		//bWithComments - false- ������ ������� - �� �����������, ����� ������� � ����������� �� �����������
		var lines = textEditor.getDocText(true);
		var methodStartRE = /^\s*�������|^\s*���������/i;
		var methodEndRE   = /^\s*������������|^\s*��������������/i;
		for (var i = curPosition; i >= 0; i--) {
			var str = lines[i];
			if (str.search(methodStartRE) != -1) {
				curPosition = i;
				break;
			}
			else if (str.search(methodEndRE) != -1) {
				break;
			}
		}

		if (bWithComments) {
			curPosition = getPrevLineMatch(lines, /(^$)|(^\s*[^ /]+)/m, curPosition - 1);
		}
		return curPosition;
	};

	this.getMethodEndLine = function (curPosition, bWithComments) {
		//bWithComments - false- ����� ������� - ��������� �� �����������, ����� ������� � ����������� ����� �����
		var lines = textEditor.getDocText(true);
		curPosition = getNextLineMatch(lines, /^\s*������������|^\s*��������������/i, curPosition);
		if (curPosition != -1) {
			curPosition = getPrevLineMatch(lines, /(^$)|(^\s*[^ /]+)/m, curPosition+1);
		}
		return curPosition;
	};

}

//��������� ��������� �� ���� �������
function checkDoc() {
	var doc = CommonScripts.GetTextDocIfOpened(0);
	mTextEditor.setDoc(doc);
	if (doc) {
		return 1;
	}
	else {
		return 0;
	}
}

function createMethod(name, isFunction) {
	if (!name) return;
}

var methodSeparator= "//******************************************************************************\r\n";
var funcDef = "������� <���>()\r\n<����>\t������� �;\r\n������������\t\/\/ <���>";
var procDef = "��������� <���>()\r\n<����>��������������\t\/\/ <���>";

function buildMethodText(methodName, methodType, methodBody) {
	var methodText;
	if (!methodType) {
		//������� ��� ������ (�������/���������)
		var svc = new ActiveXObject("Svcsvc.Service");
//		var methodType = svc.SelectValue("function\r\nprocedure", "", false);
  		methodType = svc.PopupMenu("function\r\nprocedure", 1, 0, 0);
	}
	if (methodType) {
		var methodDef = methodType == "function" ? funcDef : procDef;
		methodText = methodDef.replace(/<���>/g, methodName);
		methodText = methodText.replace(/<����>/g, methodBody);
		methodText = methodSeparator + methodText;
	}
	return methodText;
}

// ���� �������� ������, �� ������� ���������/������� � ������ ���������� ����� � ���
// ���� ������(�) �� ��������, �� ���� � ������� ������ ����� ���������/������� � ������� ��
//  ���� ����� �� ������, �� �����������
function extractMethod() {
	if (checkDoc()) {
		var lDoc = mTextEditor.getDoc();
		var insertAt = mTextAnalyzer.getMethodStartLine(lDoc.SelStartLine, true);
		if (insertAt == -1) return;
		
		var methodName = "";
		var methodBody = "";
		var methodType = "";
		
		if (lDoc.SelStartLine == lDoc.SelEndLine) {
			var selText = mTextEditor.getSelectedText();
			//���� ������� ����� �� ����� ������ (�� ��� ������),
			// �� ������� �������/��������� � ������ ���������� ������
			if (selText) {
				methodName = selText;
			}
			//���� ������ �� ��������
			// �� ������� �������/��������� � ������ ������� ������ ������-���� ������
			else {
				methodName = mTextAnalyzer.getMethodCallAtLine(lDoc.SelStartLine);
			}
			//��������� ��� ������ �� ������� ����� �����
			var t = mTextEditor.getText(lDoc.SelStartLine, 0, lDoc.SelStartLine, null);
			if (t.indexOf("=") != -1) methodType = "function";
			else methodType = "procedure";
		}
		// ����� ��� ������ ���������, ��������� ���������� ����� � ����� �����
		else {
			methodName = CommonScripts.InputBox("������� ��� ���������/�������");
			methodBody = mTextEditor.getSelectedText();
			//mTextEditor.deleteLines(lDoc.SelStartLine, mTextEditor.getSelectionHeight());
		}
		
		if (methodName) {
			var methodText = buildMethodText(methodName, methodType, methodBody);
			if (methodText) {
				if (methodBody) lDoc.Cut();

				mTextEditor.insertLines(mTextEditor.text2LinesArray(methodText), insertAt);
				var line = insertAt + 2;	//����� ��� � ����������� //*****

				//������������� ���� ������ ������
				var height = mTextAnalyzer.getMethodEndLine(line);
				lDoc.MoveCaret(line, 0, line + height, 0);
				lDoc.FormatSel();

				//������ � ������ ������
				var col = mTextEditor.getText(line, null, line, null).length - 1;
				lDoc.MoveCaret(line, col, line, col);
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
	}
	catch (e) {
		Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);
		Message(e.description, mRedErr);
		Message("������ " + SelfScript.Name + " �� ��������", mInformation);
		Scripts.UnLoad(SelfScript.Name); 		
	}
}

function testRe(text, re) {
	var matches = text.match(re);
	if (matches) {
		for (var i = 0; i < matches.length; i++) {
			Message(matches[i]);
		}
	}
	else {
		Message("������ �� �������� �� �������!");
	}
}

init(0);

var mTextEditor = new TextEditor();
var mTextAnalyzer = new TextAnalyzer(mTextEditor);
