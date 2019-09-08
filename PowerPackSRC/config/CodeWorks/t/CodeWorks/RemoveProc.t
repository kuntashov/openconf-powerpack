#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 6 }

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
$expected = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1


����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::RemoveProc("����������2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================
# ����������� ��������
$input = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1

// ����������2()
// ���������:
//	���
// ������������ ��������:
//	�����, 2
������� ����������2()
	������� 2;
������������ // ����������2

����������� = "������, ���!";
��������(�����������);
EOF
;
$expected = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1

// ����������2()
// ���������:
//	���
// ������������ ��������:
//	�����, 2

����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::RemoveProc("����������2"), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 3 ======================================================
# ����������� ���������
$input = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1

// ����������2()
// ���������:
//	���
// ������������ ��������:
//	�����, 2
������� ����������2()
	������� 2;
������������ // ����������2

����������� = "������, ���!";
��������(�����������);
EOF
;
$expected = <<EOF
����� �����������;

������� ����������1()
	������� 1;
������������ // ����������1


����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
ok(CodeWorks::RemoveProc("����������2", 1), 1);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);
