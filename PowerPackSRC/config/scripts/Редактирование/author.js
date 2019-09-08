$NAME ��������� �����������
/*===========================================================================
Copyright (c) 2004-2005 Alexander Kuntashov
=============================================================================
������:  author.js ("��������� �����������")
������:  $Revision: 1.9 $
�����:   ��������� ��������
E-mail:  kuntashov at yandex dot ru
ICQ:     338758861
�����������: 
    CommonServices.wsc Registry.wsc
��������: 
    �������� ���� �� ��������� ������������������ � �������� ����������
����������� (������� ������). �������� ������ ��������� ������� � 
�������������� ����������� ����������. ������� (���������) ������ ����������� 
�� ����� ������ (�������� �������������� ������������� ����� �������� 
������������ 1� ��� ������������ ������������ �������), �������� ����������� 
� ������� ����. ��� ����� ������� �����������; ���� ������������� � ������������
� �������� ��������. ��������� ������� ������������ � ��������� ����.

===========================================================================*/

/*===========================================================================
                                �������
===========================================================================*/


//@������: �������������� ()
//  ��������� ���������� ���� ��� ������� ������ ��������� "��������"
function ��������������()
{
    setMarker(1);
}

//@������: ������������� ()
//  �������� ���������� ���� ��� ������� ������ ��� �����������
//  � ��������� ����/������ ��������� "�������"
function �������������()
{
    setMarker(2);
}

//@������: ������������ ()
//  ������������ ���������� ����/������ � ������������� ����������
//  ������� "������"
function ������������()
{
    setMarker(3);
}

//@������: ��������� ()
//  �������� ���� ��������� ���������� �������
function ���������()
{
    OpenSettingsWindow();
}

//@������: ������������� ()
//  ��������� ����� �� ������ ������ � �����������
//  ��������������� ��������
function �������������()
{
    PasteTextFromClipboard();
}

//@������: ���������������������������������� ()
//  ��������� � ������ ��������� ��������� ������� ���� � �����
//  � ��������� � ��������� ���
function ����������������������������������()
{
    AddTimeStampAndSave();
}

/*===================================================================================
                            ������ ���. ����
=====================================================================================
    ��� ���������� ����������� ������������ ��������� "Registry.wsc"

    ��������� �������� � ������� � 
        HKCU\Software\1C\1Cv7\7.7\OpenConf\Scripts\<$NAME ����� c������> 
                
    ������������ ��������� ��������� (��� - ���������): 
        authorName - ���/������� ������, � ������ ������������ ��������� �����������,
                     ��� ������� ����������� ���������������:
                     %1CUser% - ������� ������������ 1�
                     %OSUser% - ������� ������������ (������������) �������
                
        companyName - �������� �����������
                
        dateFormat - ������ ����
        
        markerAdded     - ����������� ������ ������������ �����
        markerChanged   - ����������� ������ ����������� �����
        markerDeleted   - ����������� ������ ���������� �����
        markerEndBlock  - ����������� ������ �����
        
        oneliner - ���� �����, �� ������������ � �������� ������� ��� 
                   ������������� �����������, ���� �� �����, �� � �������� ������� 
                   ������������� ����������� ����� �������������� ���������������
                   ����������� ������ �����               
                   
        signature - ���������� ���������, ������� ������������� ����� �������,
                    ������������ ��������� �����������:
                    %Author% - ���������� �� �������� ��������� authorName
                    %Company% - ���������� �� �������� ��������� companyName
                    %Date% - ���������� �� ������� ����, ����������������� � 
                             ������������ � dateFormat

        splitter - ����������� ���� ��� ������

        doNotIndent - �� ������������ ����� ��� ���������/������ - ��� � "������ ����" ������ ��������
                      (�� ��������� ������������ ������ ����� ��, ��� � ������ ������ �����)

        doNotSignAtEnd - �� ��������� ��������� � ����������� ������ �����

        doNotCopyOldCode - �� ���������� ������ ��� ����� ����������

        addModDateTimeAtEnd - ��������� ����� ���������� ��������� � ����� �����
                             (�� ��������� ����������� � ������ �����)
        
===================================================================================*/

var __AuthorJSSettingsWindowCaption = "��������� ����������� :: ���������";
var __AuthorJSSettingsWindowPath; // ���� �� HTML-����� ���� ��������

function indent(line)
{
    var m = line.match(/^(\s*)/);
    if (m) {
        if (m[0] !== line) {
            return m[1];
        }
    }
    return '';
}

function setMarker(markerType, newCode)
{
    var s = null, doc = null;
    if (!(doc = CommonScripts.GetTextDocIfOpened())) {
        return;
    }
    with (new Settings(CommonScripts.Registry, SelfScript.Name)) {
        Load();
        s = GetData();
    }

    /* 
     *  ������������� ���������� ����������� 
     */
    
    // ������
    var AC_Marker = "";
    switch (markerType) {
        case 1:
            AC_Marker = s.markerAdded;
            break;
        case 2:
            AC_Marker = s.markerChanged;
            break;
        case 3:
            AC_Marker = s.markerDeleted;
            break;
        default:
            return; // �� ������ ���� � ��������
    }
    
    // ����������� ������� � ������ ���� ��� ������
    var AC_Splitter = s.splitter; 

    // ���� � �����
    var AC_Date = getCurrentDate(s.dateFormat);
    var AC_Time = getCurrentTime(); // ��������� %Time% �������� ��� �������������

    // �����
    var AC_Author = s.authorName.replace(/%1CUser%/i, get1CUser()).replace(/%OSUser%/i, getOSUser());

    // ���������
    var AC_Sign = s.signature.replace(/%Author%/i, AC_Author)
                .replace(/%Company%/i, s.companyName)
                .replace(/%Date%/i, AC_Date)
                .replace(/%Time%/i, AC_Time);
    
    if (((doc.SelStartLine == doc.SelEndLine) && ((markerType != 2) || s.doNotCopyOldCode) && (!newCode))) {
        // ��������� ������������� �������
        
        var line = doc.Range(doc.SelStartLine).replace(/\s*$/i, "");
        if (s.oneliner !== null) { // ���� ��� ������������� ����� ���� ������, �� ��� � ����������
            AC_Marker = s.oneliner;
        }
        doc.Range(doc.SelStartLine) = (markerType == 3?"//":"") + line + " //" + AC_Marker + AC_Sign;   
    }
    else {
        // �������� �����
    
        var selEndLine  = doc.SelEndLine - (doc.SelEndCol?0:1);     
        var block       = doc.Range(doc.SelStartLine, 0, selEndLine, doc.LineLen(selEndLine));
    
        var lines   = block.split(/\r\n/);
        var ind     = s.doNotIndent?"":indent(lines[0]);                                    
        
        // ����������� ������ �����
        block = "\r\n" + ind + "//" + AC_Marker + AC_Sign + "\r\n";
        
        // �������� ��� ������ ���� - ������������ ����
        if (markerType != 1) {  
            if ((markerType != 2) || !s.doNotCopyOldCode) {
                block += commentLines(lines, ind);      
            }
        }

        // ������ ���� - ��������� ����������� ������� � ������ ����
        if (markerType == 2) {
            // ��� ����������� ����� ������ ����������� �� ������ (XXX ����� ������� �����������?)
            if (AC_Splitter && !s.doNotCopyOldCode && (doc.SelStartLine != doc.SelEndLine)) {
                block += ind + "//" + AC_Splitter + "\r\n";
            }   
        }
            
        // ����������� ��� ������������� ���
        if (markerType != 3) {                  
            block += (newCode?newCode:lines.join("\r\n")) + "\r\n";
        }
            
        // ����������� ������ �����
        block += ind + "//" + s.markerEndBlock + (s.doNotSignAtEnd?"":AC_Sign) + "\r\n";
                    
        // ��������� ��������������� ��� � ��������     
        doc.Range(doc.SelStartLine, 0, selEndLine, doc.LineLen(selEndLine)) = block;        

        // ������������� ������ ������ �����, ���� ��� ����������, ����� - ����� ����� ����
        doc.MoveCaret(doc.SelStartLine + lines.length + 3, 0);
    }
}

function PasteTextFromClipboard(_)
{
    var doc = null, clipboard = CommonScripts.GetFromClipboard();

    if (!clipboard) {
        return;
    }
    
    if (!(doc = CommonScripts.GetTextDocIfOpened())) {
        return;
    }
    
    var marker = ((doc.SelStartLine != doc.SelEndLine)||(doc.SelStartCol != doc.SelEndCol))?2:1;
        
    setMarker(marker, clipboard);
}

function AddTimeStampAndSave(_)
{
    var s = null, doc = null;
    if (!(doc = CommonScripts.GetTextDocIfOpened())) {
        return;
    }
    with (new Settings(CommonScripts.Registry, SelfScript.Name)) {
        Load();
        s = GetData();
    }
    
    var lineNo = s.addModDateTimeAtEnd ? doc.LineCount : 0;
    var newTimeStamp = "//@ " + getCurrentDate(s.dateFormat);
    
    var line = doc.Range(lineNo);
    
    if (!line.match(/^\/\/@/)) {
        if (s.addModDateTimeAtEnd) {
            newTimeStamp = doc.Range(lineNo) + "\r\n" + newTimeStamp;
        }
        else {
            newTimeStamp += "\r\n" + doc.Range(lineNo);
        }
    }
    
    doc.Range(lineNo) = newTimeStamp;
    Windows.ActiveWnd.Document.Save();
}

function commentLines(lines, ind)
{
    var ret = "";
    for (var i=0; i<lines.length; i++) {
        ret += ind + "//" + lines[i] + "\r\n";
    }
    return ret;
}

// ����������� ������ ������� �� TEXAS :-)
function ZeroZero(num)
{
    return (num>9)?num:('0'+num);
}

function getCurrentDate(format)
{
    with (new Date) {           
        return format.replace(/yyyy/, getYear())
        .replace(/yy/, (new String(getYear())).substr(2,2))
        .replace(/dd/, ZeroZero(getDate()))
        .replace(/mm/, ZeroZero(getMonth()+1))
        .replace(/HH/, ZeroZero(getHours()))
        .replace(/MM/, ZeroZero(getMinutes()))
        .replace(/SS/, ZeroZero(getSeconds()))
    }
}

function getCurrentTime(_)
{
    with (new Date) {
        return getHours() + ':' + (getMinutes() + 1) + '.' + getSeconds();
    }
}

function get1CUser(_)
{
    return Configurator.AppProps(appUserName);
}

function getOSUser(_)
{
    return CommonScripts.WSH.ExpandEnvironmentStrings("%USERNAME%");
}

function Settings(OCReg, sname)
{       
//private:
    var sname = sname;
    var settings = {
    // ��������� �� ���������
        authorName      : "MyName",
        companyName     : "MyCompany",
        dateFormat      : "yyyy-mm-dd HH:SS:MM",
        markerAdded     : "+",
        markerChanged   : "*",
        markerDeleted   : "-",
        markerEndBlock  : "/",
        oneliner        : null,
        signature       : "%Author%@%Company%, %Date%",
        splitter        : " -------- �������� ��:",
        doNotIndent     : null,
        doNotSignAtEnd  : null,
        doNotCopyOldCode: null,
        addModDateTimeAtEnd: null
    }
    
//public:   
    this.Load = function ()
    {
        var val = null, srk = OCReg.ScriptRootKey(sname);
        
        for (var a in settings) {
            if ((val = OCReg.Param(srk, a)) !== null) {
                settings[a] = val;
            }
        }       
    }
    this.Fill = function (f)
    {
        var ix = 2;
        for (var a in settings) {           
            if (settings[a] !== null) {
                f[a].value = settings[a];
            }
        }   
        f.oneliner.disabled = f.rbOneliner[0].checked = (settings.oneliner === null);
        f.rbOneliner[1].checked = !f.rbOneliner[0].checked;
        if (settings.authorName.match(new RegExp("^%1CUser%$", "i"))) {
            ix = 0;
        }
        if (settings.authorName.match(new RegExp("^%OSUser%$", "i"))) {
            ix = 1;
        }
        f.rbAuthorType[ix].checked = true;
        f.authorName.disabled = (ix !== 2);

        f.doNotCopyOldCode_     .checked = settings["doNotCopyOldCode"];
        f.doNotSignAtEnd_       .checked = settings["doNotSignAtEnd"];
        f.doNotIndent_          .checked = settings["doNotIndent"];
        f.addModDateTimeAtEnd_  .checked = settings["addModDateTimeAtEnd"];

    }
    this.Read = function (f) 
    {
        for (var a in settings) {
            settings[a] = f[a].value;
        }
        if (f.rbOneliner[0].checked) {
            settings.oneliner = null;
        }
    }   
    this.Save = function ()
    {
        var srk = OCReg.ScriptRootKey(sname);
        for (var a in settings) {
            if (settings[a] !== null) {
                OCReg.Param(srk, a) = settings[a];
            }
        }   
        if (settings.oneliner === null) {
            OCReg.DeleteParam(srk, "oneliner");
        }
    }
    this.GetData = function ()
    {
        return settings;
    }
}

function initForm(f) 
{
    with (new Settings(CommonScripts.Registry, SelfScript.Name)) {
        Load();
        Fill(f);
    }
}

function saveSettings(f)
{
    with (new Settings(CommonScripts.Registry, SelfScript.Name)) {
        Read(f);
        Save();
    }
    closeWindow();
}

function closeWindow(_)
{
    Windows.ActiveWnd.Close();
    with (CommonScripts.FSO) {
        if (FileExists(__AuthorJSSettingsWindowPath)) {
            DeleteFile(__AuthorJSSettingsWindowPath, true);
        }
    }
    reloadSelf();
}

function reloadSelf(_)
{
    Scripts.Reload(SelfScript.Name);
}

function handleDCEvent(_)
{
    with (CommonScripts) {  
        if (Registry == null) {
            Error(GetLastError() + "\n" + "���������� ������� ��� ������ � ��������!");
            return;
        }
    }           
    try {       
        var h = {init:initForm, ok:saveSettings, cancel:closeWindow, free:reloadSelf};
        SettingsWindow.Document.Script.SetHandlers(h);      
    }
    catch (e) {
        Message("������ ������������� ���� ��������", mRedErr);
        Message(e.description, mRedErr);
    }
}

function OpenSettingsWindow(_)
{
    var f,wnd = null;
    if (wnd = CommonScripts.FindWindow(__AuthorJSSettingsWindowCaption)) {
        Windows.ActiveWnd = wnd;
    } 
    else {
        with (CommonScripts.FSO) {                  
            __AuthorJSSettingsWindowPath = "D:\\Temp\\author.html";//BuildPath(GetSpecialFolder(2), GetTempName());
            f = CreateTextFile(__AuthorJSSettingsWindowPath, true);
            f.Write(getHtml());
            f.Close();          
        }       
        wnd = OpenOleForm("Shell.Explorer", __AuthorJSSettingsWindowCaption);
        SelfScript.AddNamedItem("SettingsWindow", wnd, false);
        eval('function SettingsWindow::DocumentComplete(d,u){ return handleDCEvent()}');
        SettingsWindow.Navigate2(__AuthorJSSettingsWindowPath);
    }   
}

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

/* ������������� ������������� � ������� html2scr.js */
function getHtml(_)
{
    return ''
    +"<html>\r\n"
    +"  <head>\r\n"
    +"  <title>��������� ������� &quot;��������� �����������&quot;</title>\r\n"
    +"<style type=\"text/css\">\r\n"
    +"body {\r\n"
    +"  font-family         : sans-serif;\r\n"
    +"  font-size           : small;\r\n"
    +"  background-color    : #CCCCCC;\r\n"
    +"}\r\n"
    +"</style>\r\n"
    +"<script language=\"JScript\">\r\n"
    +"var handlers\r\n"
    +"function SetHandlers(h) { handlers = h; handlers.init(settings) }\r\n"
    +"function setCheckboxValue(c) { settings[c.name.replace(/_$/,'')].value = c.checked?\"1\":\"\"; }\r\n"
    +"</script>\r\n"
    +"  \r\n"
    +"</head>\r\n"
    +"<body onUnload=\"handlers.free()\" >\r\n"
    +"<form name=\"settings\">\r\n"
    +"<center>\r\n"
    +"  <table width=\"750 pt\" border=\"1 pt\">\r\n"
    +"      <tr><td>\r\n"
    +"\r\n"
    +"<table width=\"100%\" height=\"100%\">\r\n"
    +"  <tr><td><b>��������� ���������</b></td></tr>\r\n"
    +"  <tr align=\"top\"><td>\r\n"
    +"  <table width=\"100%\">\r\n"
    +"      <tr><td colspan=\"2\">�����:</td></tr>\r\n"
    +"      <tr><td>&nbsp;</td><td><input type=\"radio\" name=\"rbAuthorType\" value=\"1\" onClick=\"authorName.value='%1CUser%';authorName.disabled=true\"/>��� ������������ 1�</td></tr>\r\n"
    +"      <tr><td>&nbsp;</td><td><input type=\"radio\" name=\"rbAuthorType\" value=\"2\" onClick=\"authorName.value='%OSUser%';authorName.disabled=true\"/>��� ������������ Windows</td></tr>\r\n"
    +"      <tr><td>&nbsp;</td><td><input type=\"radio\" name=\"rbAuthorType\" value=\"3\" onClick=\"authorName.disabled=false\" checked/>�������:&nbsp;<input name=\"authorName\" type=\"text\" value=\"\"/></td></tr>\r\n"
    +"  </table>\r\n"
    +"  </td></tr>\r\n"
    +"  <tr><td>\r\n"
    +"  <table>\r\n"
    +"      <tr><td>�����������:&nbsp;</td><td><input name=\"companyName\" type=\"text\" value=\"\"/></td></tr>\r\n"
    +"  </table>\r\n"
    +"  </td></tr>\r\n"
    +"  <tr><td>\r\n"
    +"  <table>\r\n"
    +"      <tr><td>������ ����:&nbsp;</td><td><input name=\"dateFormat\" type=\"text\" value=\"yyyy/mm/dd\"/></td></tr>        \r\n"
    +"  </table>\r\n"
    +"  <table width=\"100%\">\r\n"
    +"      <tr><td>������ ���������:&nbsp;</td></tr>       \r\n"
    +"      <tr><td align=\"center\"><input name=\"signature\" type=\"text\" value=\"%Author%@%Company%, %Date%\" size=\"50\"/></td></tr>\r\n"
    +"  </table>    \r\n"
    +"  </td></tr>\r\n"
    +"  </table>            \r\n"
    +"          </td>\r\n"
    +"          <td width=\"50%\" align=\"top\">\r\n"
    +"  \r\n"
    +"<table width=\"100%\" >\r\n"
    +"  <tr><td><b>������ ������������</b></td></tr>\r\n"
    +"  <tr><td>\r\n"
    +"  <table width=\"100%\">\r\n"
    +"      <tr><td colspan=\"2\">����������� ���� �������:</td></tr>\r\n"
    +"      <tr><td>&nbsp;&nbsp;&nbsp;�������� ��������:</td><td><input name=\"markerAdded\" type=\"text\" value=\"+\"/></td></tr>\r\n"
    +"      <tr><td>&nbsp;&nbsp;&nbsp;�������� �������:</td><td><input name=\"markerChanged\" type=\"text\" value=\"*\"/></td></tr>\r\n"
    +"      <tr><td>&nbsp;&nbsp;&nbsp;�������� ������:</td><td><input name=\"markerDeleted\" type=\"text\" value=\"-\"/></td></tr>\r\n"
    +"  </table>\r\n"
    +"  </td></tr>\r\n"
    +"  <tr><td>\r\n"
    +"  <table>\r\n"
    +"      <tr><td>����������� ������ �����:&nbsp;</td><td><input name=\"markerEndBlock\" type=\"text\" value=\"/\"/></td></tr>\r\n"
    +"  </table>\r\n"
    +"  </td></tr>\r\n"
    +"  <tr><td>\r\n"
    +"  <table width=\"100%\">\r\n"
    +"      <tr><td colspan=\"2\">������������ ������:</td></tr>\r\n"
    +"      <tr><td>&nbsp;</td><td><input name=\"rbOneliner\" type=\"radio\" value=\"1\" onClick=\"oneliner.disabled=true\" checked/>������������ ������ �����</td></tr>\r\n"
    +"      <tr><td>&nbsp;</td><td><input name=\"rbOneliner\" type=\"radio\" value=\"2\" onClick=\"oneliner.disabled=false\"/>������:&nbsp;<input name=\"oneliner\" type=\"text\" value=\"\"/></td></tr>\r\n"
    +"  </table>\r\n"
    +"  </td></tr>\r\n"
    +"  <tr><td>����������� ���� ��� ������:</td></tr>\r\n"
    +"  <tr><td align=\"center\"><input name=\"splitter\" type=\"text\" value=\" -------------\" size=\"40\"/></td></tr>\r\n"
    +"</table>      \r\n"
    +"              \r\n"
    +"          </td>\r\n"
    +"      </tr>\r\n"
    +"  </table><br/>\r\n"
    +"  <input name=\"doNotCopyOldCode\"    type=\"hidden\" value=\"\" />\r\n"
    +"  <input name=\"doNotSignAtEnd\"  type=\"hidden\" value=\"\" />\r\n"
    +"  <input name=\"doNotIndent\"     type=\"hidden\" value=\"\" />\r\n"
    +"  <input name=\"addModDateTimeAtEnd\" type=\"hidden\" value=\"\" />\r\n"
    +"  <table width=\"750 pt\">\r\n"
    +"      <tr><td><b>�������������� ���������</b></td></tr>\r\n"
    +"      <tr><td><input name=\"doNotCopyOldCode_\" type=\"checkbox\" onClick=\"splitter.disabled=this.checked;setCheckboxValue(this);\"/>�� ��������� ����� ���� ��� ������/��������� </td></tr>\r\n"
    +"      <tr><td><input name=\"doNotSignAtEnd_\" type=\"checkbox\" onClick=\"setCheckboxValue(this);\"/>�� ��������� ��������� � ����������� ������ �����</td></tr>\r\n"
    +"      <tr><td><input name=\"doNotIndent_\" type=\"checkbox\" onClick=\"setCheckboxValue(this);\"/>�� ������������ ������ (�� ��������� ������ ��� � ������ ������ � �����)</td></tr>\r\n"
    +"      <tr><td><input name=\"addModDateTimeAtEnd_\" type=\"checkbox\" onClick=\"setCheckboxValue(this);\"/>��������� ���� � ����� ��������� � ����� �����</td></tr>\r\n"
    +"  </table>    \r\n"
    +"  <br/>\r\n"
    +"  <input type=\"button\" name=\"btOk\" value=\"��������� ���������\" onClick=\"handlers.ok(settings)\" />&nbsp;\r\n"
    +"  <input type=\"button\" name=\"btCancel\" value=\"������\" onClick=\"handlers.cancel()\" />  \r\n"
    +"</form>   \r\n"
    +"</center>\r\n"
    +"\r\n"
    +"<a href=\"http://1c.proclub.ru/modules/mydownloads/personal.php?cid=1004&lid=4915\" target=\"_blank\" title=\"�������� �������� ������� �� ��������\">�������� �������� �������� ������� �� ��������</a><br/>\r\n"
    +"<a href=\"mailto:kuntashov@yandex.ru?subject=OC_AuthorJS:\" title=\"��������� �������� kuntashov@yandex.ru\">�������� ������ ������ �������</a></br>\r\n"
    +"<a href=\"http://openconf.itland.ru\" target=\"_blank\" title=\"�������� �������� ������� &quot;�������� ������������&quot;\">������ &quot;�������� ������������&quot;</a><br/>\r\n"
    +"</body>\r\n"
    +"</html>\r\n"
}

