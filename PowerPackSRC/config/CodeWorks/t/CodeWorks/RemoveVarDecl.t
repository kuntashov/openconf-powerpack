#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 4 }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = <<EOF
����� �������������;
EOF
;

# �������� ��� ������ �� ����� 1�, 
# ��������� ��������� ���������� �������
$expected = <<EOF
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

ok(CodeWorks::RemoveVarDecl("","","�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================

$input = <<EOF
��������� ������������()
	����� �������������;
	��������("������, ���!");
�������������� // ������������()
EOF
;
$expected = <<EOF
��������� ������������()
	��������("������, ���!");
�������������� // ������������()
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

ok(CodeWorks::RemoveVarDecl("", "������������", "�������������"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

