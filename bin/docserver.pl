#!/usr/bin/perl -w

$^W = 1;
use strict;
use RPC::PlServer;
use lib '.';
use Docserver;

package Docserver::Srv;
use vars qw( @ISA $VERSION );
@ISA = qw( RPC::PlServer );
$VERSION = $Docserver::VERSION;

my $server = new Docserver::Srv( { qw(
		debug		20
		facility	stderr
		logfile		docserver.log
		mode		single
		localport	5454 ),
		'methods' => {
			'Docserver::Srv' => {
				'NewHandle' => 1,
				'CallMethod' => 1,
				'DestroyHandle' => 1,
				'errstr' => 1,
				},
			'Docserver' => {
				'new' => 1,
				'stderr' => 1,
				'preferred_chunk_size' => 1,
				'input_file_length' => 1,
				'put' => 1,
				'convert' => 1,
				'result_length' => 1,
				'get' => 1,
				'finished' => 1,
				'errstr' => 1,
				}
			},
		'clients' => [
			{
			'mask' => '\.fi\.muni\.cz$',
			'accept' => 1,
			},
			],	
		} );
$server->Bind();

sub errstr {
	return "Server error: $Docserver::errstr";
	}