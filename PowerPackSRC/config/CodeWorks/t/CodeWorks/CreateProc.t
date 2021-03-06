#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 10, todo => [8] }

use vars qw( $input $output $expected $code );

$code = <<EOF;
��������� ������������()
	// ���-���
�������������� // ������������()
EOF

#== 1 ======================================================
# ������ ������
$input = "";
$expected = <<EOF;
��������� ������������()
	// ���-���
�������������� // ������������()
EOF
@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 2 ======================================================
# ������ ����������� ����������
$input = "����� �����1;";
$expected = <<EOF
����� �����1;
��������� ������������()
	// ���-���
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 3 ======================================================
# ��� ���� ����� ���������
$input = <<EOF
����� �����1;
��������� ������������()
	// ���-���
�������������� // ������������()
EOF
;
$expected = <<EOF
����� �����1;
��������� ������������()
	// ���-���
�������������� // ������������()
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 4 ======================================================
# ���� ��� � ������ (�� ����������!)
$input = <<EOF
��������("������, ���!");
EOF
;
$expected = <<EOF
��������� ������������()
	// ���-���
�������������� // ������������()

��������("������, ���!");
EOF
;
@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 5 ======================================================
# ���������� ���������� + ��� � ������
$input = <<EOF
����� �����������;
����������� = "������, ���!";
��������(�����������);
EOF
;
$expected = <<EOF
����� �����������;
��������� ������������()
	// ���-���
�������������� // ������������()

����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 6 ======================================================

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

��������� ������������()
	// ���-���
�������������� // ������������()

������� ����������2()
	������� 2;
������������ // ����������2

����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code, "����������2");
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 7 ======================================================
 
$input = <<EOF
����� �����������;

// ����������1()
// ���������:
//	���
// ������������ ��������:
//	�����, 1
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

// ����������1()
// ���������:
//	���
// ������������ ��������:
//	�����, 1
������� ����������1()
	������� 1;
������������ // ����������1

��������� ������������()
	// ���-���
�������������� // ������������()

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

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code, "����������2");
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 8 ======================================================
# FIXME �������� ��, ����� ��� ������� ��������� ����� ������
# !!! ����� ���� ����������� ������ ������

$input = <<EOF
����� �����������;

// ����������1()
// ���������:
//	���
// ������������ ��������:
//	�����, 1
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

// ����������1()
// ���������:
//	���
// ������������ ��������:
//	�����, 1
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

��������� ������������()
	// ���-���
�������������� // ������������()

����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code);
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

#== 9 ======================================================

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

��������� ������������()
	// ���-���
�������������� // ������������()

������� ����������2()
	������� 2;
������������ // ����������2

����������� = "������, ���!";
��������(�����������);
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code, "�����������, ����������2");
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);

@CodeWorks::lines = split(/\n/, $input, -1);
CodeWorks::CreateProc("������������", $code, "����������2, �����������");
$output = join("\n", @CodeWorks::lines);
ok($output, $expected);
