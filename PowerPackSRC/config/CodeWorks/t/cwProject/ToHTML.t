#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 3 }

use vars qw( $cwp $expected );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

$cwp->Add('InsertVarDecl', "", "", "��������������������");
$cwp->Add('CreateProc', "MyNewProc", <<PROC, "MyFunc");
Procedure MyNewProc()
	Message("Hello, World!");	
EndProcedure // MyNewProc()
PROC
$cwp->Add('RenameObject', "", "MyFunc", "MyLocalVar", "MyRenamedVar");
$cwp->Add('InsertVarDecl', "", "MyNewProc", "SomeVar");
$cwp->Add('RemoveProc', "MyProc", 1);
ok($cwp->Count(), 5);

ok($cwp->ToHTML()."\n",<<HTML);
<table class="cwProjectView">
<tr class="darkRow">
<td><input type="checkbox" name="chbAction0" /></td>
<td><a class="action" href="javascript:void(0)" onClick="editAction(0);">�������� ���������� ���������� <code>��������������������</code> � ������</a></td>
</tr>
<tr class="lightRow">
<td><input type="checkbox" name="chbAction1" /></td>
<td><a class="action" href="javascript:void(0)" onClick="editAction(1);">������� ������������ <code>MyNewProc</code> ����� ������������� <code>MyFunc</code></a></td>
</tr>
<tr class="darkRow">
<td><input type="checkbox" name="chbAction2" /></td>
<td><a class="action" href="javascript:void(0)" onClick="editAction(2);">������������� ������ �� <code>MyLocalVar</code> � <code>MyRenamedVar</code> � ������������ <code>MyFunc</code></a></td>
</tr>
<tr class="lightRow">
<td><input type="checkbox" name="chbAction3" /></td>
<td><a class="action" href="javascript:void(0)" onClick="editAction(3);">�������� ���������� ���������� <code>SomeVar</code> � ������������ <code>MyNewProc</code></a></td>
</tr>
<tr class="darkRow">
<td><input type="checkbox" name="chbAction4" /></td>
<td><a class="action" href="javascript:void(0)" onClick="editAction(4);">������� ������������ <code>MyProc</code> ������ � �������������</a></td>
</tr>
</table>
HTML
;
