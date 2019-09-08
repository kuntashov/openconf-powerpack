#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 16 }

use vars qw( $cwp $obj );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

$obj = $cwp->CreateAction('RemoveProc', 'ProcName');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������� ������������ ProcName
EOF
;
$obj = $cwp->CreateAction('RemoveProc', 'MyFunc', 1);
ok($cwp->ActionToString($obj, 1)."\n", <<EOF);
������� ������������ <code>MyFunc</code> ������ � �������������
EOF
;
$obj = $cwp->CreateAction('InsertCode', '', "It's not important by now");
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ��� "It's not important by now" � ������
EOF
;
$obj = $cwp->CreateAction('InsertCode', 'Proc', "And this is a bit longer than 25 chars");
ok($cwp->ActionToString($obj,1)."\n", <<EOF);
�������� ��� <code>And this is a bit long...</code> � ������ ������������ <code>Proc</code>
EOF
;
$obj = $cwp->CreateAction('InsertCode', 'Proc', "And this consists\nfrom several lines", 1);
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ��� "And this consists from..." � ����� ������������ Proc
EOF
;
$obj = $cwp->CreateAction('CreateProc', 'MyProc', 'It does not matter by now');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������� ������������ MyProc
EOF
;
$obj = $cwp->CreateAction('CreateProc', 'MyProc', 'It does not matter by now', 'MyFunc');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������� ������������ MyProc ����� ������������� MyFunc
EOF
;
$obj = $cwp->CreateAction('InsertVarDecl', '', '', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ���������� ���������� MyVar � ������
EOF
;
$obj = $cwp->CreateAction('InsertVarDecl', '', 'MyProc', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ���������� ���������� MyVar � ������������ MyProc
EOF
;
$obj = $cwp->CreateAction('RemoveVarDecl', '', '', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������� ���������� ���������� MyVar �� ������
EOF
;
$obj = $cwp->CreateAction('RemoveVarDecl', '', 'MyProc', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������� ���������� ���������� MyVar �� ������������ MyProc
EOF
;
$obj = $cwp->CreateAction('RenameObject', '', '', 'OldObjName', 'NewObjName');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������������� ������ �� OldObjName � NewObjName �� ���� ������
EOF
;
$obj = $cwp->CreateAction('RenameObject', '', 'MyProc', 'OldObjName', 'NewObjName');
ok($cwp->ActionToString($obj)."\n", <<EOF);
������������� ������ �� OldObjName � NewObjName � ������������ MyProc
EOF
;
$obj = $cwp->CreateAction('ReplaceCode', '', '', '#some pattern here', '//a piece of new code');
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ������/������ "#some pattern here" �� ��� "//a piece of new code" �� ���� ������
EOF
;
$obj = $cwp->CreateAction('ReplaceCode', '', 'MyFunc', '#some pattern here', '//a piece of new code');
ok($cwp->ActionToString($obj)."\n", <<EOF);
�������� ������/������ "#some pattern here" �� ��� "//a piece of new code" � ������������ MyFunc
EOF
;