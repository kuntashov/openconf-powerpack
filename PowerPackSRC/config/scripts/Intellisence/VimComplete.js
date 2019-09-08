/*===========================================================================
Copyright (c) 2004-2005 Alexander Kuntashov
=============================================================================
������:  VimComplete.js
������:  1.4
�����:   ��������� ��������
E-mail:  kuntashov at yandex dot ru
ICQ UIN: 338758861
��������: 
    ������������� ���� � ����� ��������� Vim
===========================================================================*/
/*
    ������ NextWord()
            ��������� ����� ����� ����� �� ������� � �������� ��������� ���,
        ��� ������ �� ������ �����, � ����� �� ����� ������. ����������� 
        ������ ����������. ��������� ����� ������� ��������� ��������� �� 
        ������ ��������� ������ � ��� ����� �� ����� (����� �� ��������� ������ 
        ������ ����� ����������� � ������ ������). 
        
    ������ PrevWord()
            ���� �����, ������ ����� ���� �������������� � �������� �����������.
   
    � ������������ Vim ������������ ��������� ������:
    
        Ctrl + N  ��� ���������� � ������� ������ (��������� �����, Next word)
        Ctrl + P  ��� ���������� � ������� �����  (���������� �����, Previous word)
*/

/* ==========================================================================
                                    �������
========================================================================== */

// ��������� ������������
function NextWord() // Ctrl + N
{
    completeWord();
}

// ���������� ������������
function PrevWord() // Ctrl + P
{
    completeWord(true);
}

/* ==========================================================================
                                IMPLEMENTATION                       
========================================================================== */

var CurDoc = null;
var VimComplete = new VimAutoCompletionTool;

function completeWord(lookBackward)
{
	if (!(CurDoc = CommonScripts.GetTextDocIfOpened(false, true))) {
		// ���� �� ���� �� ���������, �� �������� ���������� ������������
		// 1�'������� ������ (���� ������� ��������)
		Configurator.CancelHotKey = true;
		return;
	}
	VimComplete.completeWord(lookBackward);
	CurDoc = null;
}

/* ���������� ����� ����� ����� �� 
    �������� ��������� ������� */
function getLeftWord(doc)
{
    var cl, word = '';
    cl = doc.Range(doc.SelStartLine);
    for (var i=doc.SelStartCol-1;
            (i>=0)&&cl.charAt(i).match(/[\w�-��-�]/i);
                word = cl.charAt(i--) + word)
        ;    
    return word;
}

/* ����������� ����� ��� ��������� ���� � ������, 
�� ���������� � ����������������� ��������. 
�������� � ���, ��� ��� ������� ��������� �� ������
(����������, ������ �� ����� � �� �������� ����� ���� 
����� �� �����, � ����� ��� ���������, ����� ���� �����������
����� ���������� �� ������������� ������������, ��� ��� ������� 
����� ��������� � ��������� ���������)*/

function Line(str)
{
// private 
    var s = str;
    var words = null;     
// public
    this.reset = function () 
    {
        words = s.split(/[^\w�-�]+/);
    }
    this.assign = function (ix) 
    {
        return ((typeof(words)=="object")&&(ix>=0)&&(ix<words.length));                  
    }
    this.word = function (ix) 
    {
        if (this.assign(ix)) return words[ix];
    }
    this.count = function ()
    {
        return words.length;        
    }
	this.words = function ()
	{
		return words;
	}
	/* ���������� ������-�������� � ������������ ������� next() 
	   c ������� �������� �������������� ������� ����� (�� ��� ���, 
	   ���� �� ������ �������� undefined, ���������� ����� ������ */
	this.iterator = function (r)
	{
		var collection = this;
		return {
			collection	: collection,
			iterator	: r ? collection.count() : (-1),														 
			next		: function(reverse)
		    {        
				return this.collection.word( this.iterator += (reverse?(-1):1) );
		    }   
		}
	}
	/* ��������� ��������, �������� ������ ��, �������� �������
	   ������ ������ pattern */
    this.filter = function (pattern, unique)
    {
        var used = {};
        if (this.assign(0)) {
            var nw = new Array();
            for (var i=0; i<this.count(); i++) {
                if (this.word(i).match(pattern)) {					
                    if (unique) {
                        if (!used[this.word(i)]) {
                            used[this.word(i)] = true;
                            nw[nw.length] = this.word(i);
                        }
                    }
                    else {
                       nw[nw.length] = this.word(i);
                    }
                }
            }
            words = nw;
            return true;
        }
        return false;
    }        
   
    this.reset(); // �������������
}

/* �� [����] ��������� ����������� ��������� ��������� ��������������
������� Vim (� ��� ^P/^N), �� ����������� ������������
����������, �� ��� �� ��������� ������ - ������������� � 
������ ������� ������������ �� ����������� �� ������ 
�����������. ���� ���� ��� ����������, ���� �� ������,
��� ������ ������������ ��������� � ���������, � �� ��� �� ����. 
��� ���� ��������� ��� ����������� � ���, ��� ��� ����������
������� ����� � ������ ������� ��������� �� ������ ����������
���������� � � ��������� (�������) �����, � ���������, �� 
��������� �����, �� ��� ������ �� �� �����������, ���� ��������,
��� ������ �������� ��� ������� :-) */

function VimAutoCompletionTool(_) // orgy citadel :-)
{
// private

    var srcDocPath;	// ���� �� ��������� ��������� (������������ ��� ������������� ����������)
    var srcLine;	// �������� ������ ���������
    var srcCol;		// ������ ������� � ������ ����� �������� ������
    var srcWord;	// �������� ����� (������� �������� ���������)
    var lastWord;	// ��������� �������������� � ����������� �����
    var curLineIx;	// ������ ������� ������ (�� ������� ������� ������������)
    var words;		// ������ ����-������������ ������� ������
    
    var backwardSearch;	// �������� ����� (�� ��������� ����� ������, "������")
    var pattern;		// ������ (���������� ���������), ����������� ������������ ��������� ����� 
    
    var counter;	// ������� ������������
    var total;		// ����� ����� ������������

// public

    /* ��������� (��)������������� �������, ���� ��� ���������� */
    this.setup = function (lookBackward)
    {
        var word = getLeftWord(CurDoc);		
        if (this.isNewLoop(CurDoc, word)) { // ���������������			
            srcDocPath	= CurDoc.Path;            
            srcLine		= CurDoc.SelStartLine;
            srcCol		= CurDoc.SelStartCol - word.length;            
            srcWord		= word;            
            lastWord	= word; // ����� ��������� ������� ������ �����������
            curLineIx	= srcLine;            
            pattern		= new RegExp("^" + word, "i");     
            // �������� ������ ������������ ������� � �������� ������
            words = this.parseLine(lookBackward ? this.leftPart() : this.rightPart());  
            // ��������
            counter = 0;
            total = null;
        }               
        backwardSearch = lookBackward;
    }
	/* ������� ������������� ���������� ����������������� 
        ���������� ������ ������� VimAutoCompletionTool */
    this.isNewLoop = function (doc, word)
    {
        return !(words 
			&& (srcDocPath	== doc.Path) 
            && (srcLine		== doc.SelStartLine) 
            && (lastWord	== word)
			&& (srcCol		== (doc.SelStartCol-word.length)));         
    }
	/* ���������, �� ������� �� ������ ������ �� ���������� ������� */
    this.assign = function (lIx)
    {
        return (CurDoc&&(0<=lIx)&&(lIx<CurDoc.LineCount));
    }
	/* ����� ��������� ������������ � ����������� ��� �� ����� ��������� ����� */
    this.completeWord = function (lookBackward)
    {       
		this.setup(lookBackward);
        while (true) {  
            var word = words.next(lookBackward);
            if (word) {
                this.complete(word);
                return;
            }
            words = this.nextLine();
        }
    }
	/* ������ � ���������� ������ �������������� ���� ��� 
		��������� �� ������� ������ */
    this.nextLine = function ()
    {                   
        curLineIx += (backwardSearch ? -1 : 1); 
		if (backwardSearch) {
			if (curLineIx < 0) {
				curLineIx = CurDoc.LineCount-1;
			}
		} 
		else {
            if (curLineIx == CurDoc.LineCount) {
                curLineIx = 0;
            }
         }
        return this.parseLine(this.curLine());               
    }
	/* "���������" ���������� � �������� ��������� ������ �� �����
		� ��������� �� � ����������� � ��������, ������� ���������
		���������� ������������ ��� ��������� ����� */
    this.parseLine = function (srcLine)
    {
        var w = new Line(srcLine);
        w.filter(pattern, true);       
        return w.iterator(backwardSearch);
    }
	/* ��������� ����������� ���������� ������������ ������ ��������� ����� */
    this.complete = function (word)
    {        
        CurDoc.Range(srcLine) = this.leftPart() + word + this.rightPart();
        CurDoc.MoveCaret(srcLine, srcCol+word.length);
            
        lastWord = word;

        counter += backwardSearch ? -1 : 1;          
        if ((curLineIx == srcLine)&&(lastWord == srcWord)) {
            if ((!total)&&counter) {                                
                total = Math.abs(counter) - 1;                     
            }
            counter = 0;
        }       

        Status(counter 
            ? "C�����������: "+Math.abs(counter).toString()+(total?" �� " + total:"")
            : (total ? "�������� �����" : "������ �� ������")); 
    } 
    /* ���������� ������� ������ */
    this.curLine = function ()
    {
        var str = "";
        if (this.assign(curLineIx)) {
            /* ��������� �������� ������ � ��� ��������� ��������,
                �� "��������" ��������, ��������� �� ����� �������� ����� */
            if (curLineIx == srcLine) {
                str = this.leftPart() + srcWord + this.rightPart();           
            }
            else {
                str = CurDoc.Range(curLineIx);
            }           
        }
        return str;
    }
    /* ����� �������� ������, ���������� �������� �����; 
        ���� �������� ����� �� ���������� */
    this.leftPart = function ()
    {
        return CurDoc.Range(srcLine, 0, srcLine, srcCol);
    }
    /* ������ �������� ������, ���������� �������� �����;
        ���� �������� ����� �� ���������� */
    this.rightPart = function ()
    {      
        return CurDoc.Range(
            srcLine, srcCol+lastWord.length,
            srcLine, CurDoc.LineLen(srcLine));
    }
}

/*
 * ��������� ������������� ������� 
 */
function Init(_) // ��������� ��������, ����� ��������� �� �������� � �������
{    
    try {
        var c = null;
        if (!(c = new ActiveXObject("OpenConf.CommonServices"))) {
            throw(true); // �������� ����������
        }        
        c.SetConfig(Configurator);
        SelfScript.AddNamedItem("CommonScripts", c, false);
    }
    catch (e) {
        Message("�� ���� ������� ������ OpenConf.CommonServices", mRedErr);    
        Message("������ " & SelfScript.Name & " �� ��������", mInformation);
        Scripts.UnLoad(SelfScript.Name);
    }
}

Init(); // ��� �������� ������� ��������� ��� �������������
