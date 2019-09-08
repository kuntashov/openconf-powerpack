#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 1 }

use vars qw( $input $output $expected );

$input = <<EOF
Перем Перем1;

//******************************************
Процедура МояПроцедура()
	// бла-бла
КонецПроцедуры // МояПроцедура()
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

CodeWorks::ProcessModule("МояПроцедура");


ok($CodeWorks::proc_start, 3);
