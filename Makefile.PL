
print STDERR <<EOF;

	Please read the README file. It's not likely that the software
	will run of box for you. At least you'll have to modify the
	configuration to suit your conditions.

EOF

sleep 5;

use ExtUtils::MakeMaker;
WriteMakefile(
	'NAME' =>		'Docserver',
	'VERSION_FROM' =>	'lib/Docclient.pm',
	'EXE_FILES' =>		[ 'bin/docclient.pl', 'bin/docserver.pl' ],
	'dist' =>		{ COMPRESS => 'gzip -9f', SUFFIX => 'gz',
		POSTOP => 'mv $(DISTNAME)-$(VERSION).tar.gz ../' },
	);
