#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 16 }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = "";
$expected = "Перем МояПеременная;";
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================

$input = <<EOF
Перем МояПеременная1;
EOF
;
$expected = <<EOF
Перем МояПеременная1;
Перем МояПеременная2;
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","МояПеременная2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 3 ======================================================

$input = <<EOF
Процедура МояПроцедура()
КонецПроцедуры // МояПроцедура()
EOF
;
$expected = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная;
КонецПроцедуры // МояПроцедура()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","МояПроцедура","МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 4 ======================================================

$input = <<EOF
Процедура МояПроцедура()
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;
$expected = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная;
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","МояПроцедура","МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 5 ======================================================

$input = <<EOF
Сообщить("Привет, Мир!");
EOF
;
$expected = <<EOF
Перем МояПеременная;
Сообщить("Привет, Мир!");
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 6 ======================================================

$input = <<EOF
Перем МояПеременная1;
Сообщить("Привет, Мир!");
EOF
;
$expected = <<EOF
Перем МояПеременная1;
Перем МояПеременная2;
Сообщить("Привет, Мир!");
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","МояПеременная2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 7 ======================================================

$input = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная1;
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;
$expected = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная1;
	Перем МояПеременная2;
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","МояПроцедура","МояПеременная2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 8 ======================================================

$input = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная;
КонецПроцедуры // МояПроцедура()
EOF
;
$expected = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная;
КонецПроцедуры // МояПроцедура()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","МояПроцедура","МояПеременная"), 0);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);
