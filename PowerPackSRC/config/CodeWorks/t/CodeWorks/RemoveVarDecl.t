#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 4 }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = <<EOF
Перем МояПеременная;
EOF
;

# исходный код модуля на языке 1С, 
# ожидаемый результат применения шаблона
$expected = <<EOF
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

ok(CodeWorks::RemoveVarDecl("","","МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================

$input = <<EOF
Процедура МояПроцедура()
	Перем МояПеременная;
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;
$expected = <<EOF
Процедура МояПроцедура()
	Сообщить("Привет, Мир!");
КонецПроцедуры // МояПроцедура()
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

ok(CodeWorks::RemoveVarDecl("", "МояПроцедура", "МояПеременная"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

