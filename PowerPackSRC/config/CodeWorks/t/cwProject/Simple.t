#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 13 }

use vars qw( $cwp $tmp );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

$cwp->Add('RemoveProc', "MyProc", 1);
ok($cwp->Count(), 1);
ok($cwp->Action(0)->{Type},			'RemoveProc');
ok($cwp->Action(0)->{ProcName},		'MyProc');
ok($cwp->Action(0)->{WithComments}, 1);

$cwp->Add('RenameObject', "", "MyFunc", "MyLocalVar", "MyRanamedVar");
ok($cwp->Count(), 2);

$cwp->Add('CreateProc', "MyNewProc", <<PROC, "MyFunc");
Procedure MyNewProc()
	Message("Hello, World!");	
EndProcedure // MyNewProc()
PROC
ok($cwp->Count(), 3);

$tmp = $cwp->CreateAction('InsertVarDecl', "", "MyNewProc", "SomeVar");
$cwp->Add($tmp);
ok($cwp->Count(), 4);

ok($cwp->Del(2), 1);
ok($cwp->Count(), 3);

# нет такого действия
ok($cwp->CreateAction('ThisActionWeDontKnow'), undef);
# индекс за границами 
ok($cwp->Action(666), undef);
ok($cwp->Del(999), 0);
