
package Docserver::Config;
%Docserver::Config::Config =
	(
	'port'		=>	2324,
	### 'port'		=>	5455,
	'clients'	=>	[
					{
					'mask' => '^127\.0\.0\.1$',
					'accept' => 1,
					},
					{
					'mask' => '\.fi\.muni\.cz$',
					'accept' => 1,
					},
				],

	'tmp_dir'	=>	'C:\\tmp\\docserver.tmp',
	'ChunkSize'	=>	128 * 512,
	'ps'		=>	'Generic PostScript Printer on FILE:',
	### 'ps'		=> 	'Adobe Generic PS on FILE:',
	'ps1'		=>	'Adobe Generic PS1 on FILE:',
	'pidfile'	=>	'docserver.pid',
	'logfile'	=>	'docserver.log',
	);

