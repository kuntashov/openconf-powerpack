#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 9 }

use vars qw( $cwp $obj );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

##################### RemoveProc
$obj = $cwp->CreateAction('RemoveProc', 'ProcName');
ok($cwp->ActionToPerl($obj), <<EOF);
CodeWorks::RemoveProc("ProcName", 0);
EOF
;
##################### InsertCode
$obj = $cwp->CreateAction('InsertCode', '', <<EOF, 1);
Procedure MyProc()
	// code here
EndProcdeure // MyProc()
EOF
;
ok($cwp->ActionToPerl($obj), <<EOCODE);
CodeWorks::InsertCode("", <<EOF, 1);
Procedure MyProc()
	// code here
EndProcdeure // MyProc()
EOF
;
EOCODE
;

##################### CreateProc
$obj = $cwp->CreateAction('CreateProc', 'MyProc', <<EOF);
Procedure MyProc()
	// code here
EndProcdeure // MyProc()
EOF
ok($cwp->ActionToPerl($obj), <<EOCODE);
CodeWorks::CreateProc("MyProc", <<EOF, "");
Procedure MyProc()
	// code here
EndProcdeure // MyProc()
EOF
;
EOCODE
;

##################### InsertVarDecl/RemoveVarDecl
$obj = $cwp->CreateAction('InsertVarDecl', '', '', 'MyVar');
ok($cwp->ActionToPerl($obj), <<EOF);
CodeWorks::InsertVarDecl("", "", "MyVar");
EOF
;
$obj = $cwp->CreateAction('RemoveVarDecl', '', 'MyProc', 'MyVar');
ok($cwp->ActionToPerl($obj), <<EOF);
CodeWorks::RemoveVarDecl("", "MyProc", "MyVar");
EOF
;

##################### RenameObject
$obj = $cwp->CreateAction('RenameObject', '', '', 'OldObjName', 'NewObjName');
ok($cwp->ActionToPerl($obj), <<EOF);
CodeWorks::RenameObject("", "", "OldObjName", "NewObjName");
EOF
;
$obj = $cwp->CreateAction('RenameObject', '', 'MyProc', 'OldObjName', 'NewObjName');
ok($cwp->ActionToPerl($obj), <<EOF);
CodeWorks::RenameObject("", "MyProc", "OldObjName", "NewObjName");
EOF
;

##################### ReplaceCode
$obj = $cwp->CreateAction('ReplaceCode', '', 'MyFunc', 'A = B;', 'A = B + 1;');
ok($cwp->ActionToPerl($obj), <<EOFCODE);
CodeWorks::ReplaceCode("", "MyFunc", q(A = B;), <<EOF);
A = B + 1;
EOF
;
EOFCODE
;