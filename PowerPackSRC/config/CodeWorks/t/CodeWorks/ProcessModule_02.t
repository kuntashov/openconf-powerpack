#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 1 }

use vars qw( $input $output $expected );

$input = <<EOF
����� �����1;

//******************************************
��������� ������������()
	// ���-���
�������������� // ������������()
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

CodeWorks::ProcessModule("������������");


ok($CodeWorks::proc_start, 3);
