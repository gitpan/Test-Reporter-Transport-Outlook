package Test::Reporter::Transport::Outlook;

=head1 NAME

Test::Reporter::Transport::Outlook - MS Outlook as transport for Test::Reporter

=cut

use strict;
use warnings;

require 5.008_008;

use base 'Test::Reporter::Transport';
use Mail::Outlook;
use Params::Validate qw(:types validate_pos);
use Devel::AssertOS qw(MSWin32);

=head1 SYNOPSIS

	use warnings;
	use strict;
	use Test::Reporter;
	use Test::Reporter::Transport::Outlook;

	my $reporter = Test::Reporter->new(
		grade        => 'fail',
		distribution => 'Mail-Freshmeat-1.20',
		from         => 'whoever@wherever.net (Whoever Wherever)',
		comments     => 'output of a failed make test goes here...',
		via          => 'CPANPLUS X.Y.Z',
	);

	my $sender = Test::Reporter::Transport::Outlook->new(0);

	$sender->send($reporter);

=head1 DESCRIPTION

If you want to test Perl modules for CPAN but is located in a very restricted environment, without HTTP or SMTP access
outside the network and can count only with an Outlook and Exchange server to send emails, this module is for you.

This module will use C<Win32::OLE> module with MS Outlook to create a new email message. This message will contain
all information regarding the test report and I<would> be ready to be sent I<if> new security patches of Outlook 
didn't forbid that.

As an alternative, the message can be saved on the "Drafts" (or whatever is called in your computer) or can be displayed
so you will only need to click (or review before clicking) the report. Default behavior is to save the email so you can
avoid having lot's of new email messages being opened if you're testing real hard new distributions.

Of course, this module expects that you're in a MS Windows machine with MS Outlook installed. This module will not work
in other operational systems.

=head1 USAGE

See L<Test::Reporter> and L<Test::Reporter::Transport> for general usage information.

=cut

our $VERSION = '0.01';

=head2 Transport Arguments

    $report->transport_args( 1 );

See C<new> method to understand this parameter.

=head1 METHODS

These methods are only for internal use by Test::Reporter.

=head2 new

    my $sender = Test::Reporter::Transport::Outlook->new( 0 ); 

The C<new> method is the object constructor.

There is an optional parameter that is related if the email will be saved in the Drafts folder (the default behavior)
or displayed to the user executing this module. The valid values are true (1) or false (0) respectivally. Anything 
else will generate an exception.

To save the message, nothing needs to be passed. To display the message, pass 0 as an argument.

The message will be saved or displayed upon invocation of the method C<send>.

=cut

sub new {

    my $class = shift;

    my @params =
      validate_pos( @_, { type => SCALAR, regex => qr/^1|0$/, default => 1 } );

    my $self =
      { _outlook => Mail::Outlook->new(), _is_draft => shift(@params) };

    return bless $self, $class;
}

=head2 get_outlook

Returns the C<Mail::Outlook> object that represents the MS Outlook application.

=cut

sub get_outlook {

    my $self = shift;

    return $self->{_outlook};

}

=head2 send

Display or save the message in the Outlook, depends on the values defined during object creation. See the method
C<new> for more details.

Expects as a parameter a C<Test::Reporter> object. Returns true in case of success or false in failure (1 or 0);

=cut

sub send {

    validate_pos(
        @_,
        { type => OBJECT, isa => 'Test::Reporter::Transport::Outlook' },
        { type => OBJECT, isa => 'Test::Reporter' }
    );

    my ( $self, $report ) = @_;

    my $message = $self->get_outlook()->create();

    $message->To( $report->address() );
    $message->Subject( $report->subject() );
    $message->Body( $report->report() );

    if ( $self->{_is_draft} ) {

        return $message->save();

    }
    else {

        return $message->display();

    }

}

1;

__END__

=head1 CONFIGURATION

C<Test::Reporter::Transport::Outlook> can be configured with C<CPAN::Reporter> as described in 
L<CPAN::Reporter::Config>.

In the I<config.ini> file, include the line

	transport=Outlook

to enable using of C<Test::Reporter::Transport::Outlook> as the method of transport for reports. If you want to create
new messages like drafts, nothing else needs to be done. But if you want the messages to pop up before sending them, 
then it's necessary to disable the default behaviour by using the line

	transport=Outlook 0

in the config.ini file.

=head1 CAVEATS

Even by saving the messages as drafts, sending lot's of reports by clicking "Send" message by message can be a pain. Due
to the security patches on Microsoft Outlook, the only alternative is to use the Redemption DLL to avoid the patches
applied by Microsoft and send the messages directly.

C<Test::Reporter::Transport::Outlook> uses L<Mail::Outlook|Mail::Outlook> to create and send emails messages with 
Outlook. This issue is better documented there.

There is a good change that Redemption DLL may be included in L<Mail::Outlook|Mail::Outlook> distribution and once this
is done, then it will be possible to send reports of tests without user intervention.

Another issue is the customized header (I<X-Reported-Via>) used by L<CPAN::Reporter|CPAN::Reporter> emails to help the 
CPAN testers identify which transport was used to send the received report. While sending this header is not a mandatory 
feature, do not sending it makes the life of CPAN testers more difficult to identify issues. Once again, there is not 
way to set this header by using only Outlook. To be able to send such header, it's necessary:

=over

=item 1.
Redemption DLL using

=item 2.
The Exchange server which Outlook will use to send the message must allow the specified header to be sent.

=back

Even if the Redemption DLL is used, there is no guarantee that the message will have the specified header by Outlook
side. You will need to check it with your Exchange server administrator to enable that.

=head1 SEE ALSO

=over

=item *
L<Win32::OLE>

=item *
L<Mail::Outlook>

=item *
L<Test::Reporter::Transport>

=item *
L<CPAN::Reporter::Config>

=back

=head1 AUTHOR

 Alceu Rodrigues de Freitas Junior <arfreitas@cpan.org>.

=head1 COPYRIGHT

 Copyright (C) 2009 Alceu Rodrigues de Freitas Junior

 All rights reserved.

=head1 LICENSE

This program is free software; you may redistribute it and/or modify it under
the same terms as Perl itself.

=cut

