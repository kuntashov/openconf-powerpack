/*=============================================================================

	Copyright (c) 2005 OpenConf Community	<http://openconf.itland.ru>
	Copyright (c) 2005 Alexander Kuntashov	<kuntashov@yandex.ru>

	Code Works Helper - �������� ������-�������� ��� ������� ��� ��������
	������ ���� ������� ������ (CodeIns.pl, CodeWorks.pm).

	$Id: $

	������:  
		��������� �������� aka a13x <kuntashov@yandex.ru> icq#338758861  

=============================================================================*/

var baseFrame = null;

var Handlers = 
{
	onClose			 	: onClose,
	onAddClick		 	: onAddClick,
	onDelClick       	: onDelClick,
	onUpClick		 	: onUpClick,
	onDownClick		 	: onDownClick,
	onDlgSaveClick	 	: onDlgSaveClick,
	onDlgCancelClick 	: onDlgCancelClick,
	onOpenProjectClick	: onOpenProjectClick,
	onSaveProjectClick	: onSaveProjectClick,
	editAction			: editAction
}

function enableMenu(flag)
{
	var menuItems = baseFrame.children.fmMain.children.mainMenu.children;
	menuItems['btnAdd'].disabled	= !flag;
	menuItems['btnDel'].disabled	= !flag;
	menuItems['btnUp'].disabled		= !flag;
	menuItems['btnDown'].disabled	= !flag;
	menuItems['btnSave'].disabled	= !flag;
	menuItems['btnOpen'].disabled	= !flag;
}

function showFrame(id)
{
	var main	= baseFrame.children.fmMain;
	var dialog	= baseFrame.children.fmDialog;	
	var allDlgs	= baseFrame.children.allDialogs;

	dialog.style.display = 'none';

	if (id && (id != "fmMain")) {		
		// �������� ���� �������� ����
		enableMenu(false);
		dialog.children['dlgBody'].innerHTML = allDlgs.children['dlg'+id].innerHTML;
		dialog.style.display = 'block';
		return;
	}

	enableMenu(true);
}

function onAddClick(_)
{
	var cwOps = {
		"�������� ���"						: 'InsertCode',		 
		"������������� ������"				: 'RenameObject',
		"�������� ������/������"			: 'ReplaceCode',	 
		"������� ���������� ����������"		: 'InsertVarDecl',
		"������� ���������� ����������"		: 'RemoveVarDecl', 
		"������� ������������"				: 'CreateProc',		 
		"������� ������������"				: 'RemoveProc'	 
	}
	
	var menu = "";
	for (var op in cwOps) {
		menu += op + "\n";
	}
	
	var ret = SvcSvc.PopupMenu(menu);
	if (ret&&cwOps[ret]) {
		showFrame(cwOps[ret]);
	}
}

function editAction(ix)
{
	var action = cwProject.Action(ix);
	if (!action) return;

	showFrame(action.Type);

	var inp = baseFrame.children['fmDialog'].children['dlgBody'].children;
	inp.actionIndex.value = ix;

	switch (action.Type) {
		case 'RemoveProc' :
			inp.p1.value 	= Action.ProcName;
			inp.p2.checked	= Action.WithComments;
			break;
		case 'InsertCode' :
			inp.p1.value	= action.ProcName;
			inp.p2.value	= action.Code;
			inp.p3.checked	= action.AtEnd;
			break;
		case 'CreateProc' :
			inp.p1.value = action.ProcName;
			inp.p2.value = action.ProcText;
			inp.p3.value = action.BeforeProc;
			break;
		case 'InsertVarDecl' : 
		case 'RemoveVarDecl' :
			inp.p1.value = action.ObjName;
			inp.p2.value = action.ProcName;
			inp.p3.value = action.VarName;
			break;
		case 'RenameObject' :
			inp.p1.value = action.ObjName;
			inp.p2.value = action.ProcName;
			inp.p3.value = action.OldName;
			inp.p4.value = action.NewName;
			break;
		case 'ReplaceCode' :
			inp.p1.value = action.ObjName;
			inp.p2.value = action.ProcName;
			inp.p3.value = action.OldCode;
			inp.p4.value = action.NewCode;			
			break;
		default :
			return;
	}	
}

function onSaveProjectClick(_)
{
	var filter = "����� �������� CodeWorks (*.pl)|*.pl|��� �����|*";
	var path = SvcSvc.SelectFile(true, IBDir + "��������������.pl", filter, false); 
	if (!path) return;
	if (!cwProject.SaveToFile(path)) {
		// TODO � ����� ���������
		Message("������ ���������� ����� �������!", mRedErr);
		return;
	}
}

function onOpenProjectClick(_)
{
	Message("���� �� �����������", mInformation);
}

function updateProjectView(_)
{
	var cnt = baseFrame.children.fmMain.children.projectContainer;
	cnt.innerHTML = cwProject.ToHTML();
}

function onDelClick(_)
{
	// TODO ������ �� ������������� ��������!
	var objsToDelete = [];
	var trs = baseFrame.children.fmMain.children.projectContainer.children[0].children[0].children;
	for (var i=0; i<trs.length; i++) {
		var checkbox = trs[i].children[0].children[0];
		if (!checkbox.checked) continue;
		var m = checkbox.name.match(/^chbAction(\d+)$/);
		if (!m) continue;
		objsToDelete[objsToDelete.length] = cwProject.Action(m[1]);		
	}
	for (var i=0; i<objsToDelete.length; i++) {
		cwProject.Del(objsToDelete[i]);
	}
	updateProjectView();
}

function onUpClick(_)
{
	Message("���� �� �����������", mInformation);
}

function onDownClick(_)
{
	Message("���� �� �����������", mInformation);
}

function onDlgSaveClick(_)
{
	var inp = baseFrame.children.fmDialog.children.dlgBody.children;
		
	var p = function (ix) { 
		var id = 'p' + ix;
		if(typeof(inp[id])=='object') {
			return inp[id].type=="checkbox" ? inp[id].checked : inp[id].value;
		}
	}

	var action = cwProject.CreateAction(inp['type'].value, p(1), p(2), p(3), p(4));	

	if (!action) {
		// TODO ���������� ��������� �� ������� ����� CommonServices 
		Message("������ ���������� ��������", mRedErr);
		return;
	}	

	var index = inp['actionIndex'].value;
	if (index) {
		if (!cwProject.CheckIndex(index)) {
			// �� ������ ������, ���� �� ������ ������ ���� � ��������
			Message("������ �������� ������� ������� �� �������!");
			return;
		}
		cwProject.Action(index) = action;
	}
	else {
		cwProject.Add(action);
	}

	updateProjectView();	
	showFrame('fmMain');		
}

function onDlgCancelClick(_)
{
	showFrame('fmMain');
}

function onClose(_)
{
	Scripts.Unload(SelfScript.Name);
}

function onDocumentComplete(d, u)
{
	try {				
		baseFrame = HtmlWindow.Document.Script.SetHandlers(Handlers);
		Windows.ActiveWnd.Maximized = true;
	}
	catch (e) {
		Message("������ ������������� ���� Code Works Helper", mRedErr);
		Message(e.description, mRedErr);
	}
}

function Open(_)
{
	var wnd = OpenOleForm("Shell.Explorer", "Code Works Helper");
	SelfScript.AddNamedItem("HtmlWindow", wnd, false);
    eval('function HtmlWindow::DocumentComplete(d,u){ return onDocumentComplete(d, u) }');
	HtmlWindow.Navigate2(BinDir+"config\\CodeWorks\\html\\cwHelper.htm");	
}

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

CreateObjectOrDie("SvcSvc.Service", "SvcSvc");
CreateObjectOrDie("OpenConf.CodeWorksProject", "cwProject");

Open();