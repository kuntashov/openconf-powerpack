#!perl -w

# "�����������" �������� ������

use strict;
use Test::Harness qw(runtests);

my @tests = grep !/Template.t$/, map(glob( "$_/*.t" ), @ARGV ? @ARGV : qw(CodeWorks cwProject));
runtests(@tests);
