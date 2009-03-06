#!perl

use warnings;
use strict;
use Test::More tests => 5;
use Test::Reporter::Transport::Outlook;
use Test::Reporter;
use Test::Exception;

# for testing exceptions, this should be available in most Perl versions
use CGI;

my $transporter = Test::Reporter::Transport::Outlook->new();

isa_ok( $transporter,                'Test::Reporter::Transport::Outlook' );
isa_ok( $transporter->get_outlook(), 'Mail::Outlook' );

my $reporter = Test::Reporter->new(
    grade        => 'fail',
    distribution => 'Mail-Freshmeat-1.20',
    from         => 'whoever@wherever.net (Whoever Wherever)',
    comments     => 'output of a failed make test goes here...',
    via          => 'CPANPLUS X.Y.Z',
);

my $cgi = CGI->new();

dies_ok { $transporter->send('nkni') };
dies_ok { $transporter->send($cgi) };
is( $transporter->send($reporter),
    1,
    'send method returns true if the parameter is a Test::Transporter object' );

#TODO: {
#
#    local $TODO =
#      'Mail::Outlook should provide methods to check a saved message';
#
#    ok(
#        $transporter->get_outlook()->check_draft(),
#        'message was saved in the Drafts folder in Outlook'
#    );
#
#}

