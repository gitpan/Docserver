#!/usr/bin/perl -w

$^W = 1;
use strict;

use ExtUtils::MakeMaker;

my $run_server;
my $remote_server = '';

if (open CONFIG, "testcfg") {
	while (<CONFIG>) {
		s/\s*$//;
		my ($key, $value) = split /\s*:\s*/, $_, 2;
		if ($key eq 'run_server') {
			$run_server = $value;
		} elsif ($key eq 'remote_server') {
			$remote_server = $value;
		}
	}
	close CONFIG;
}

if ($^O eq 'MSWin32') {
	$run_server = prompt "I can see that this is a Windows machine\nDo you want to run docserver tests locally? [y/n]",
		((not defined $run_server) ? 'y' : $run_server);
	$run_server = ( $run_server =~ /^[yY]/ ? 'y' : 'n' );
	if ($run_server eq 'y') {
		print "Good, I will run docserver on this machine and test it.\n";
	}
} else {
	$run_server ='n';
	print "This is not a Windows machine, we won't attempt to run docserver here.\n";
}

if ($run_server eq 'n') {
	$remote_server = prompt "If you want to run tests against some other machine, tell me the name and port\nof the machine. [machine.domain.com:port]", $remote_server;

	if ($remote_server ne '') {
		while (not $remote_server =~ /^.+:\d+$/) {
			$remote_server = prompt "The format is [machine.domain.com:port], [q] to quit:", $remote_server;
			if ($remote_server eq '' or $remote_server eq 'q') {
				$remote_server = '';
				last;
			}
		}
	}

	if ($remote_server ne '') {
		print <<'EOF';

	To run tests against remote machine, you have to have docserver
	running on it. Go to the remote machine and set the port number
	and client mask in lib/Docserver/Config.pm appropriately. Then
	run docserver on the remote machine. After that come back and
	continue with testing this client. Also check that versions
	of Storable on both machines match.

EOF
		sleep 2;
	}
}

open CONFIG, "> testcfg";
print CONFIG "run_server: $run_server\n";
print CONFIG "remote_server: $remote_server\n" if $remote_server ne '';
close CONFIG;

if ($run_server eq 'y') {
	print <<EOF;

	Please read the README file. It's not likely that the software will
	run of box for you. At least you should modify the configuration
	of the server and client in files lib/Docclient/Config.pm and
	lib/Docserver/Config.pm.

EOF
	sleep 2;
}

WriteMakefile(
	'NAME' =>		'Docserver',
	'VERSION_FROM' =>	'lib/Docserver.pm',
	'EXE_FILES' =>		[ 'bin/docclient.pl', 'bin/docserver.pl' ],
	'dist' =>		{ COMPRESS => 'gzip -9f', SUFFIX => 'gz',
		POSTOP => 'mv $(DISTNAME)-$(VERSION).tar.gz ../' },
	);

if ($^O eq 'MSWin32' and ($remote_server ne '' or $run_server eq 'y')) {
	print <<EOF;

	If you do not have make working on your platform, try running
	perl do_test.pl instead of make ; make test to run the tests.

EOF
}

