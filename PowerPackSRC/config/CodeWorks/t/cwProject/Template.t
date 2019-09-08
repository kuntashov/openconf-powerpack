#!perl -w

use strict;
use Test;
use Win32::OLE;

BEGIN { plan tests => 1 }

use vars qw( $cwp );

$cwp = Win32::OLE->new("OpenConf.CodeWorksProject") or die;

ok($cwp->Count(), 0);

