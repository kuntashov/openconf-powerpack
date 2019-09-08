#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 16 }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = "";
$expected = "����� �������������;";
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================

$input = <<EOF
����� �������������1;
EOF
;
$expected = <<EOF
����� �������������1;
����� �������������2;
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","�������������2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 3 ======================================================

$input = <<EOF
��������� ������������()
�������������� // ������������()
EOF
;
$expected = <<EOF
��������� ������������()
	����� �������������;
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","������������","�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 4 ======================================================

$input = <<EOF
��������� ������������()
	��������("������, ���!");
�������������� // ������������()
EOF
;
$expected = <<EOF
��������� ������������()
	����� �������������;
	��������("������, ���!");
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","������������","�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 5 ======================================================

$input = <<EOF
��������("������, ���!");
EOF
;
$expected = <<EOF
����� �������������;
��������("������, ���!");
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 6 ======================================================

$input = <<EOF
����� �������������1;
��������("������, ���!");
EOF
;
$expected = <<EOF
����� �������������1;
����� �������������2;
��������("������, ���!");
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","","�������������2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 7 ======================================================

$input = <<EOF
��������� ������������()
	����� �������������1;
	��������("������, ���!");
�������������� // ������������()
EOF
;
$expected = <<EOF
��������� ������������()
	����� �������������1;
	����� �������������2;
	��������("������, ���!");
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","������������","�������������2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 8 ======================================================

$input = <<EOF
��������� ������������()
	����� �������������;
�������������� // ������������()
EOF
;
$expected = <<EOF
��������� ������������()
	����� �������������;
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::InsertVarDecl("","������������","�������������"), 0);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);
