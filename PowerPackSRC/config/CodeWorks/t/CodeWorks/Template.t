#!perl -w

use strict;
use Test;
use lib qw(../../ ../);
use CodeWorks;

BEGIN { plan tests => 1 }

use vars qw( $input $output $expected );

#== 1 ======================================================

$input = <<EOF
EOF
;

$expected = <<EOF
EOF
;

@CodeWorks::lines = split(/\n/, $input, -1);

# код, который мы тестируем и бла-бла-бла

$output = join("\n", @CodeWorks::lines);

ok($output, $expected);

