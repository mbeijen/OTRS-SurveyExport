# --
# Kernel/Modules/AgentSurveyExport.pm - a survey export module
# Copyright (C) 2013 tuxwerk OHG, http://tuxwerk.de/
# --
# This software comes with ABSOLUTELY NO WARRANTY. For details, see
# the enclosed file COPYING for license information (AGPL). If you
# did not receive this file, see http://www.gnu.org/licenses/agpl.txt.
# --

package Kernel::Modules::AgentSurveyExport;

use strict;
use warnings;

sub new {
    my ( $Type, %Param ) = @_;

    # allocate new hash for object
    my $Self = {};
    bless( $Self, $Type );

    # get common objects
    %{$Self} = %Param;

    return $Self;
}

sub Run {
    my ( $Self, %Param ) = @_;

    my $SurveyObject = $Kernel::OM->Get('Kernel::System::Survey');
    my $CSVObject = $Kernel::OM->Get('Kernel::System::CSV');
    my $HTMLUtilsObject = $Kernel::OM->Get('Kernel::System::HTMLUtils');
    my $LayoutObject = $Kernel::OM->Get('Kernel::Output::HTML::Layout');

    my $SurveyID = $Kernel::OM->Get('Kernel::System::Web::Request')->GetParam( Param => "SurveyID" );

    my @Questions = $SurveyObject->QuestionList(
        SurveyID => $SurveyID
    );
    my @CSVHead;
    my @CSVData;

    push @CSVHead, 'Send time';
    push @CSVHead, 'Vote time';
    push @CSVHead, 'Queue';
    push @CSVHead, 'Ticket';
    push @CSVHead, 'Ticket owner';
    foreach(@Questions) {
	push @CSVHead, $_->{Question};
    }

    my @List = $SurveyObject->VoteList( SurveyID => $SurveyID );
    # Sendezeit/SendTime, Abstimmungszeit/VoteTime
    for my $Vote (@List) {
	my @Data;
	my %Ticket = $Kernel::OM->Get('Kernel::System::Ticket')->TicketGet( TicketID => $Vote->{TicketID} );
	push @Data, $Vote->{SendTime};
	push @Data, $Vote->{VoteTime};
	push @Data, $Ticket{Queue};
	push @Data, "#".$Ticket{TicketNumber};
	push @Data, $Ticket{Owner};

	for my $Question (@Questions) {
	    my $Answer = "";
	    my @Answers = $SurveyObject->VoteGet(
		RequestID  => $Vote->{RequestID},
		QuestionID => $Question->{QuestionID},
            );
	    if ( $Question->{Type} eq 'Radio' || $Question->{Type} eq 'Checkbox' ) {
		for my $Row (@Answers) {
                    my %AnswerText = $SurveyObject->AnswerGet( AnswerID => $Row->{VoteValue} );
                    $Answer .= $AnswerText{Answer};
                }
	    }
	    elsif ( $Question->{Type} eq 'YesNo' || $Question->{Type} eq 'Textarea' ) {
		$Answer = $Answers[0]->{VoteValue};
		# clean html
                if ( $Question->{Type} eq 'Textarea' && $Answer ) {
                    $Answer =~ s{\A\$html\/text\$\s(.*)}{$1}xms;
                    $Answer = $HTMLUtilsObject->ToAscii( String => $Answer );
		    # make excel linebreak in cell work
		    $Answer =~ s/\n/\r/g;
                }
	    }
	    push @Data, $Answer;
	}

	push @CSVData, \@Data;
    }

    # get Separator from language file
    my $UserCSVSeparator = $LayoutObject->{LanguageObject}->{Separator};

    if ( $Kernel::OM->Get('Kernel::Config')->Get('PreferencesGroups')->{CSVSeparator}->{Active} ) {
        my %UserData = $Self->{UserObject}->GetUserData( UserID => $Self->{UserID} );
        $UserCSVSeparator = $UserData{UserCSVSeparator};
    }

    my $CSV = $CSVObject->Array2CSV(
	Head      => \@CSVHead,
	Data      => \@CSVData,
	Separator => $UserCSVSeparator,
    );

    return $LayoutObject->Attachment(
	Filename    => "survey.csv",
	ContentType => "text/csv; charset=utf8",
	Content     => $CSV,
    );
}

1;
