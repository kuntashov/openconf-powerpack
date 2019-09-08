#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 3 }

use vars qw( $cwp $expected );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

$cwp->Add('InsertVarDecl', "", "", "√лобальна€ѕеременна€");
$cwp->Add('CreateProc', "MyNewProc", <<PROC, "MyFunc");
Procedure MyNewProc()
	Message("Hello, World!");	
EndProcedure // MyNewProc()
PROC
$cwp->Add('RenameObject', "", "MyFunc", "MyLocalVar", "MyRenamedVar");
$cwp->Add('InsertVarDecl', "", "MyNewProc", "SomeVar");
$cwp->Add('RemoveProc', "MyProc", 1);
ok($cwp->Count(), 5);

ok($cwp->ToPerl(), <<EOCODE);
CodeWorks::InsertVarDecl("", "", "√лобальна€ѕеременна€");

CodeWorks::CreateProc("MyNewProc", <<EOF, "MyFunc");
Procedure MyNewProc()
	Message("Hello, World!");	
EndProcedure // MyNewProc()
EOF
;

CodeWorks::RenameObject("", "MyFunc", "MyLocalVar", "MyRenamedVar");

CodeWorks::InsertVarDecl("", "MyNewProc", "SomeVar");

CodeWorks::RemoveProc("MyProc", 1);

EOCODE
;