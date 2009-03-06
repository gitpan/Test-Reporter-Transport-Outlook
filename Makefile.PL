use inc::Module::Install;

name 'Test-Reporter-Transport-Outlook';
all_from 'lib/Test/Reporter/Transport/Outlook.pm';
author 'Alceu Rodrigues de Freitas Junior <arfreitas@cpan.org>';
license 'perl';

requires 'Mail::Outlook'    => '0.15';
requires 'Params::Validate' => '0.91';
requires 'Devel::CheckOS'   => '1.50';
requires 'Test::Reporter'   => '1.5203';

build_requires 'Test::More';

test_requires 'Test::POD'           => '1.22';
test_requires 'Test::Pod::Coverage' => '1.08';
test_requires 'Test::Exception'     => '0.27';

WriteAll;

