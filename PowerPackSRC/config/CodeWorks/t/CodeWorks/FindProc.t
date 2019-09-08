#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 16, todo => [14, 15, 16] }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1

������� ����������2()
	������� 2;
������������ // ����������2

����������� = "������, ���!";
��������(�����������);
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::FindProc("����������1"), 1);
ok($CodeWorks::proc_start, 2);
ok($CodeWorks::proc_end, 4);

ok($CodeWorks::proc_body, "������� ����������1()\n\t������� 1;\n������������ // ����������1");

ok(CodeWorks::FindProc("����������2"), 1);
ok($CodeWorks::proc_start, 6);
ok($CodeWorks::proc_end, 8);

#== 2 ======================================================

$input = <<EOF
����� �����������;

������� ����������() �����
��������� ������������ �����

������� ����������()
	������� 1;
������������ // ����������()

��������� ������������()
	������� 2;
�������������� // ������������()

����������� = "������, ���!";
��������(�����������);
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::FindProc("����������"), 1);
ok($CodeWorks::proc_start, 5);
ok($CodeWorks::proc_end, 7);

ok(CodeWorks::FindProc("������������"), 1);
ok($CodeWorks::proc_start, 9);
ok($CodeWorks::proc_end, 11);

#== 3 ======================================================
# TODO ����� �� ������ ������������� ����� ��������?
# ����������� � �������� � C++ �������. ��������,
# � ����� ������� ������ ������� ����������� � ������������.
$input = <<EOF
����� ��������;

������� ����������() ������� 1; ������������ // ����������()

��������� ������������()
	������� 2;
�������������� // ������������()

EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::FindProc("����������"), 1);
ok($CodeWorks::proc_start, 2);
ok($CodeWorks::proc_end, 2);
              	