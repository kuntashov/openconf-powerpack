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
Удалить подпрограмму ProcName
EOF
;
$obj = $cwp->CreateAction('RemoveProc', 'MyFunc', 1);
ok($cwp->ActionToString($obj, 1)."\n", <<EOF);
Удалить подпрограмму <code>MyFunc</code> вместе с комментариями
EOF
;
$obj = $cwp->CreateAction('InsertCode', '', "It's not important by now");
ok($cwp->ActionToString($obj)."\n", <<EOF);
Вставить код "It's not important by now" в модуль
EOF
;
$obj = $cwp->CreateAction('InsertCode', 'Proc', "And this is a bit longer than 25 chars");
ok($cwp->ActionToString($obj,1)."\n", <<EOF);
Вставить код <code>And this is a bit long...</code> в начало подпрограммы <code>Proc</code>
EOF
;
$obj = $cwp->CreateAction('InsertCode', 'Proc', "And this consists\nfrom several lines", 1);
ok($cwp->ActionToString($obj)."\n", <<EOF);
Вставить код "And this consists from..." в конец подпрограммы Proc
EOF
;
$obj = $cwp->CreateAction('CreateProc', 'MyProc', 'It does not matter by now');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Создать подпрограмму MyProc
EOF
;
$obj = $cwp->CreateAction('CreateProc', 'MyProc', 'It does not matter by now', 'MyFunc');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Создать подпрограмму MyProc перед подпрограммой MyFunc
EOF
;
$obj = $cwp->CreateAction('InsertVarDecl', '', '', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Добавить объявление переменной MyVar в модуль
EOF
;
$obj = $cwp->CreateAction('InsertVarDecl', '', 'MyProc', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Добавить объявление переменной MyVar в подпрограмму MyProc
EOF
;
$obj = $cwp->CreateAction('RemoveVarDecl', '', '', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Удалить объявление переменной MyVar из модуля
EOF
;
$obj = $cwp->CreateAction('RemoveVarDecl', '', 'MyProc', 'MyVar');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Удалить объявление переменной MyVar из подпрограммы MyProc
EOF
;
$obj = $cwp->CreateAction('RenameObject', '', '', 'OldObjName', 'NewObjName');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Переименовать объект из OldObjName в NewObjName во всем модуле
EOF
;
$obj = $cwp->CreateAction('RenameObject', '', 'MyProc', 'OldObjName', 'NewObjName');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Переименовать объект из OldObjName в NewObjName в подпрограмме MyProc
EOF
;
$obj = $cwp->CreateAction('ReplaceCode', '', '', '#some pattern here', '//a piece of new code');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Заменить строку/шаблон "#some pattern here" на код "//a piece of new code" во всем модуле
EOF
;
$obj = $cwp->CreateAction('ReplaceCode', '', 'MyFunc', '#some pattern here', '//a piece of new code');
ok($cwp->ActionToString($obj)."\n", <<EOF);
Заменить строку/шаблон "#some pattern here" на код "//a piece of new code" в подпрограмме MyFunc
EOF
;