#!/Users/aprice/perl5/perlbrew/perls/perl-5.32.0/bin/perl
use strict;
use Class::Struct;
use Date::Simple ('date', 'today');
use Text::CSV;
#use Tie::Handle::CSV;
use Data::Dumper::Simple;

use constant SUNDAY => 0;
use constant MONDAY => 1;
use constant TUESDAY => 2;
use constant WEDNESDAY => 3;
use constant THURSDAY => 4;
use constant FRIDAY => 5;
use constant SATURDAY => 6;

# Define custom data structures:
struct Lecture => {
	day_of_week => '$', # 0 means Sunday
	start_time => '$',
	location  => '$',
	duration => '$',
};

struct Tutorial => {
	day_of_week => '$', # 0 means Sunday
	start_time => '$',
	location  => '$',
	duration => '$',
};


# Initialize with the term's timetable:
# Times in the format hh:mm AM/PM
my $LHA = Lecture->new();
$LHA->day_of_week(WEDNESDAY); # use symbolic constants declared above
$LHA->start_time("9:30 AM");
$LHA->location("HSB-240");
$LHA->duration("1:00");

my $LHB = Lecture->new();
$LHB->day_of_week(MONDAY);
$LHB->start_time("9:30 AM");
$LHB->location("HSB-240");
$LHB->duration("1:00");

my $LHC = Lecture->new();
$LHC->day_of_week(TUESDAY);
$LHC->start_time("1:30 PM");
$LHC->location("HSB-240");
$LHC->duration("1:00");

# Set parameters for term from UWO calendar:
my $term_start_date = Date::Simple->new('2021-09-08');
my $term_end_date = Date::Simple->new('2021-12-08');
my $reading_week_start_date = Date::Simple->new('2021-11-01'); # Monday of Reading Week
my $reading_week_end_date = $reading_week_start_date + 4;
my $tutorial_start_date = Date::Simple->new('2021-09-13'); # Second week of term
my $thanksgiving_holiday = Date::Simple->new('2021-10-11'); # Holiday

# set some standard parameters:
#my $lesson_entry_type = "Lecture";
my $lesson_entry_type = "Class section - Lecture";

# Load topics from CSV file:
my @lesson_topics;
my $csv_in = Text::CSV->new ( { binary => 1, auto_diag => 1 } )  # should set binary attribute.
                 or die "Cannot use CSV: ".Text::CSV->error_diag ();
open my $fh, "<:encoding(utf8)", "topics.csv" or die "topics.csv: $!";
$csv_in->column_names ($csv_in->getline ($fh)); # use header
my @headings = $csv_in->column_names;
#print "Columns found in data file: $headings[0], $headings[1]\n";

print("\* List of imported topics:\n");
while ( my $row = $csv_in->getline_hr( $fh ) ) {
     #warn Dumper($row);
     printf "\tTopic: %-32s Reference: %s\n", $row->{Topic}, $row->{Reference};
     push @lesson_topics, $row;
}
print("Import of lesson topics complete: " . scalar @lesson_topics . " lessons found.\n\n");
close $fh;

# Establish a filehandle for the OWL-compatible .CSV output file:
my @owl_headings = qw(Title Description Date Start Duration Type Location Frequency Interval Ends Repeat TestCustomProperty Required);
my $csv_out = Text::CSV->new ( { binary => 1, auto_diag => 1 } )  # should set binary attribute.
                 or die "Cannot use CSV: ".Text::CSV->error_diag ();
$csv_out->eol ("\r\n");
open $fh, ">:encoding(utf8)", "scheduled_lessons.csv" or die "scheduled_lessons.csv: $!";
$csv_out->print ($fh, $_) for \@owl_headings;

print("\* Populating calendar:\n");
# Beginning at start of term, walk through each day and populate sessions as the days match:
my $topic_index = 0;
for (my $date_index = $term_start_date; $date_index <= $term_end_date; $date_index++) {
	#print("Processing $date_index.\n");

	# Check if current date falls on reading week, weekend, or Good Friday:
	if (
		(($date_index ge $reading_week_start_date) & ($date_index le $reading_week_end_date)) # reading week
		|
		($date_index->day_of_week eq 0)|($date_index->day_of_week eq 6) # weekend
		|
		($date_index eq $thanksgiving_holiday)) { # holiday
		#print("Skipping reading week, weekend or holiday date: $date_index.\n");
	}
	else {
		my @lesson;
		# First populate lecture hours in order:
		if ($date_index->day_of_week eq $LHA->day_of_week) {
	 		if (($topic_index+1) > (scalar @lesson_topics)) { # Check if number of lessons required exceeds the number imported:
				die "Not enough topics (" . scalar @lesson_topics . " found) to populate in calendar (at least " . ($topic_index+1) . " required as of $date_index).\n";
			}
	 		# add lecture topic to this date
	 		print("Lesson \"$lesson_topics[$topic_index]->{Topic}\" associated with LHA on $date_index at " . $LHA->start_time . ".\n");
	 		@lesson = ("Lesson " . ($topic_index+1) . " - " . $lesson_topics[$topic_index]->{Topic}, $lesson_topics[$topic_index]->{Reference}, $date_index->format("%m/%d/%Y"), $LHA->start_time, $LHA->duration, $lesson_entry_type, $LHA->location);
	 		$csv_out->print ($fh, $_) for \@lesson;
			$topic_index++;
		}
		if ($date_index->day_of_week eq $LHB->day_of_week) {
			if (($topic_index+1) > (scalar @lesson_topics)) { # Check if number of lessons required exceeds the number imported:
				die "Not enough topics (" . scalar @lesson_topics . " found) to populate in calendar (at least " . ($topic_index+1) . " required as of $date_index).\n";
			}
	 		# add lecture topic to this date
	 		print("Lesson \"$lesson_topics[$topic_index]->{Topic}\" associated with LHB on $date_index at " . $LHB->start_time . ".\n");
	 		@lesson = ("Lesson " . ($topic_index+1) . " - " . $lesson_topics[$topic_index]->{Topic}, $lesson_topics[$topic_index]->{Reference}, $date_index->format("%m/%d/%Y"), $LHB->start_time, $LHB->duration, $lesson_entry_type, $LHB->location);
	 		$csv_out->print ($fh, $_) for \@lesson;
			$topic_index++;
		}
		if ($date_index->day_of_week eq $LHC->day_of_week) {
	 		if (($topic_index+1) > (scalar @lesson_topics)) { # Check if number of lessons required exceeds the number imported:
				die "Not enough topics (" . scalar @lesson_topics . " found) to populate in calendar (at least " . ($topic_index+1) . " required as of $date_index).\n";
			}
	 		# add lecture topic to this date
	 		print("Lesson \"$lesson_topics[$topic_index]->{Topic}\" associated with LHC on $date_index at " . $LHC->start_time . ".\n");
	 		@lesson = ("Lesson " . ($topic_index+1) . " - " . $lesson_topics[$topic_index]->{Topic}, $lesson_topics[$topic_index]->{Reference}, $date_index->format("%m/%d/%Y"), $LHC->start_time, $LHC->duration, $lesson_entry_type, $LHC->location);
	 		$csv_out->print ($fh, $_) for \@lesson;
			$topic_index++;
		}

		# Next populate tutorials and quizzes in order:
		# (Left for future)
		# Next populate design project deliverables in order:
		# (Left for future)
	}
}
# close the output file:
close $fh or die "lessons.csv: $!";

print("* Total topics populated: " . $topic_index . " scheduled of " . (scalar @lesson_topics) . " imported.\n"); # Note that the topic index was incremented after the last population, so no correction from 0 required
if (($topic_index) < (scalar @lesson_topics)) {
	warn "Warning: there are unscheduled topics.\n";
}
