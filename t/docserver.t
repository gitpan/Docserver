#!/usr/bin/perl -w

use strict;

if ($^O =~ /win/i)
	{
	print "1..1\n";
	eval 'use Docserver';
	print "ok 1\n";
	}
else
	{
	print STDERR "\nThis doesn't look like Windows, we won't try Docserver.\n";
	print "1..0\n";
	}
