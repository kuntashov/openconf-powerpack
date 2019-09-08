function include(fileName)
{
    with ( new ActiveXObject("Scripting.FileSystemObject") ) {        
        return OpenTextFile(GetAbsolutePathName(fileName)).ReadAll();    
    }
}

function  print(str) {
    WScript.Echo(str);
}

function assign(result) {
    return result ? 'Ok' : 'Failed';
}

var frm = new ActiveXObject("OpenConf.DialogObject");

var stream = '{"Dialogs", {"Frame",'
+'{"-11","0","0","0","400","0","0","0","204","1","2","1","34","MS Sans Serif","193","151"," �����","","","0","","1","1","6","25","-1","0","0",'
+'{"1", {"����1","1"},{"����2","0"}},"1","1"}},{"Controls",{"������","BUTTON","1342177291","69","90","39","13","0","0","4152","","","","0","U","0","0","0","0","0","","","","0","-11","0","0","0","0","0","0","0","0","0","0","0","0","MS Sans Serif","-1","-1","0","����1","{""0"",""0""}"}},{"Cnt_Ver","10001"}}';

/**
 * property Stream as String, Read/Write
 *  Member of OpenConf.DialogObject
 *      �������� ��� Dialog'� (DialogStream)
 */

frm.Stream = stream;


/**
 * property Frame as OpenConfDialogFrame, ReadOnly
 *  Member of OpenConf.DialogObject
 *      ����� �������
 */

/**
 * property Property(ix) as String, Read/Write
 *  Member of OpenConfDialogFrame
 * �������� Frame'� � ������� ix (1<=ix<=30)
 *  ��� ����������� ���� �����, � ��� ����� � ��������, �������� ��� ������!!!      
 *  (����, ��� JScript ��� � �� ����� �����) �� ����,                                                                                         
 *      (frm.Frame.Property(0) == 100)
 *  ������� (�.�. ����� ���������� ����� ��������� ���������� �����), �
 *      (frm.Frame.Property(0) === 100)
 *  ������ false (����)
 */
frm.Frame.Property(1) = 100;
// ��� ����������� ���� ����� �������� ��� ������!!!! (����, ��� JScript ��� � �� ����� �����) 
// �� ����, 
// (frm.Frame.Property(0) == 100) - ������� (�.�. ����� ���������� ����� ��������� ���������� �����)
// (frm.Frame.Property(0) === 100) - ����!
print(assign(frm.Frame.Property(1)=="100"));

// �������� 29 � 30
print(assign(frm.Frame.Property(29)=="1"));
frm.Frame.Property(29) = 2;
print(assign(frm.Frame.Property(29)==2));

print(assign(frm.Frame.Property(30)=="1"));
frm.Frame.Property(30) = 3;
print(assign(frm.Frame.Property(30)==3));

/**
 * property ActiveLayer as Integer, Read/Write
 *  Member of OpenConfDialogFrame
 *      ���������� ����� ��������� ���� ����� � 
 *       ������ �����, ������� � 0
 */
print(assign(frm.Frame.ActiveLayer==1));
frm.Frame.ActiveLayer = 0;
print(assign(frm.Frame.ActiveLayer===0));

// ������� ��������� ������ ������� ��� ������� ���������� ������
// �� �������� �������� �������
frm.Frame.ActiveLayer = 10;
print(assign(frm.Frame.ActiveLayer===0));


/**
 * property Layer(ix) as OpenConfDialogFrameLayer, ReadOnly
 *  Member of OpenConfDialogFrame
 *      ���� ����� � ������� ix. ��������� ����� ������������ � 0. 
 */

/**
 * property Name as String, Read/Write
 *  Member of OpenConfDialogFrameLayer 
 *      ������������� ����
 */

print(assign(frm.Frame.Layer(0).Name == "����1"));
print(assign(frm.Frame.Layer(1).Name == "����2"));

frm.Frame.Layer(0).Name = "Layer1";
print(assign(frm.Frame.Layer(0).Name == "Layer1"));

frm.Frame.Layer(1).Name = "Layer2";
print(assign(frm.Frame.Layer(1).Name == "Layer2"));


/**
 * property Visible as Boolean, Read/Write
 *  Member of OpenConfDialogFrameLayer
 *      ��������� ����
 */

print(assign(frm.Frame.Layer(0).Visible === true));
print(assign(frm.Frame.Layer(1).Visible === false));

frm.Frame.Layer(0).Visible = false;
print(assign(frm.Frame.Layer(0).Visible === false));

frm.Frame.Layer(1).Visible = true;
print(assign(frm.Frame.Layer(1).Visible === true));

/**
 * property Browser as OpenConfDialogBrowser, ReadOnly
 *  Member of OpenConfDialogObject
 */

print(assign(frm.Browser === null));

/**
 * property Controls as OpenConfDialogControls, ReadOnly
 *  Member of OpenConfDialogObject
 *      ��������� ��������� ���������� �����
 */

/**
 * property Size as Integer, ReadOnly
 *  Member of OpenConfDialogControls
 *      ���������� ��������� ����������
 */

print(assign(frm.Controls.Count === 1));

/**
 * property Item(ix) as OpenConfDialogControl, ReadOnly
 *  Member of OpenConfDialogControls
 *      ������� ���������� � ���������� ������� ix (0<=ix<Controls.Count)
 */

// ������ �� ���������! - ������ ������� undefined
print(assign(frm.Controls.Item(100) === undefined));

/**
 * property Property(ix) as String, Read/Write
 *  Member of OpenConfDialogControl
 *      �������� � ������� ix �������� ���������� (1<=ix<=43)
 */

print(assign(frm.Controls.Item(0).Property(1) == "������"));

frm.Controls.Item(0).Property(1) = "Ok";
print(assign(frm.Controls.Item(0).Property(1) == "Ok"));

// ������ �� ���������
print(assign(frm.Controls.Item(0).Property(999) === undefined));

//var ctl = frm.Controls.Item(0);

//eval(include("./ctlProps.js"));

var ctl = frm.CreateControl("BUTTON");
frm.Controls.Add(ctl);

print(frm.Stream);