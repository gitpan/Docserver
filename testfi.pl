#!/usr/bin/perl -w

print STDERR "Beware: this test will only work on FI!\n";


my $expected = <<EOF;
Krtek je ná¹ kamarád. Pojïme si hrát.

Krtku krtku, vystrè rù¾ky.
EOF

use ExtUtils::testlib;
my $libs = join " -I", '', @INC;
my $got = `$^X $libs blib/script/docclient.pl t/test.doc`;

if ($expected ne $got)
	{ print "Expected:\n${expected}Got:\n${got}not "; }
print "ok 1\n";

