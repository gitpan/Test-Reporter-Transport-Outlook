#!perl -T

use Test::More tests => 1;

BEGIN {
	use_ok( 'Test::Reporter::Transport::Outlook' );
}

diag( "Testing Test::Reporter::Transport::Outlook $Test::Reporter::Transport::Outlook::VERSION, Perl $], $^X" );
