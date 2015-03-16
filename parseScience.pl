#!/usr/bin/env perl
# parseScience.pl by Amory Meltzer
# Parse a KSP persistent.sfs file, report science information
# Sun represented as Kerbol
# Leftover science in red, candidates for manual cleanup in green

## Fix for 0.90 beta by adding in ALL BIOMES
## Ignores KSC/LaunchPad/Runway/etc. "biomes", asteroids
## Option to pull KSC stuff in/out of Kerbin
## Cleanup data/test hashes, the order of the data is unintuitive
## Can you do srfsplashed in every biome on other planets with water?

### Add support for:
## KSC biomes
## Asteroids

### FIXES TODO
## SCANsat spacing in -t output (capitalization fix)
## Flag SCANsat on/off, or auto-detect??


use strict;
use warnings;
use diagnostics;

use Getopt::Std;
use Excel::Writer::XLSX;

# Parse command line options
my %opts = ();
getopts('atspncu:hH', \%opts);

if ($opts{h} || $opts{H}) {
  usage(); exit;
}


### FILE DEFINITIONS
my $outfile = 'scienceToDo.xlsx';
my $csvOut = 'scienceToDo.csv';

my $scidef = 'ScienceDefs.cfg';
my $pers = 'persistent.sfs';

# Change this to match the location of your KSP install
if ($opts{u}) {
  my $path = '/Applications/KSP_osx';
  $scidef = "$path/GameData/Squad/Resources/ScienceDefs.cfg";
  $pers = "$path/saves/$opts{u}/persistent.sfs";
}

# Test files for existance
if (! -e $scidef) {
  print "No ScienceDefs.cfg file found at $scidef\n";
  exit;
}
if (! -e $pers) {
  print "No persistent.sfs file found at $pers\n";
  exit;
}

### GLOBAL VARIABLES
my %dataMatrix;			# Hold errything
my %reco;			# Separate hash for craft recovery
my %scan;			# Separate hash for SCANsat
my %sbvData;			# Hold sbv values from END data

# Access reverse-engineered caps for recovery missions.  SubOrbited and
# Orbited are messed up - the default values from Kerbin are inverted
# elsewhere. ;;;;;; ##### FIXME TODO
my %recoCaps = (
		Flew => 6,
		FlewBy => 7.2,
		SubOrbited => 9.6,
		Orbited => 12,
		Surfaced => 18
	       );
# All SCANsat caps are 20
my %scanCaps = (
		AltimetryLoRes => 20,
		BiomeAnomaly => 20,
		AltimetryHiRes => 10
	       );
my %workVars;		      # Hash of arrays to hold worksheets, current row
my %spobData;		      # Hold data on science per spob
my %testData;		      # Hold data on science per test

# ScienceDefs.cfg variables
my (
    @testdef,			# Basic test names
    @sitmask,			# Where test is valid
    @biomask,			# Where biomes for test matter
    @atmo,			# Check if atmosphere required or not
    @dataScale,			# dataScale, same as dsc in persistent.sfs
    @scienceCap			# Base experiment cap, multiplied by sbv
   );

# persistent.sfs variables
my (
    @title,			# Long, displayed name
    @dsc,			# Data scale
    @scv,			# Percent left to research
    @sbv,			# Base balue multiplier to reach cap
    @sci,			# Science researched so far
    @cap			# Max science
   );

# Store details from split id
my @pieces;
my (
    @test,			# Which test
    @spob,			# Which planet/moon
    @where,			# What activity
    @biome			# What biome
   );

my @planets = qw (Kerbin Mun Minmus Kerbol Moho Eve Gilly Duna Ike Dres
		  Jool Laythe Vall Tylo Bop Pol Eeloo);
my $planetCount = scalar @planets - 1; # Use this a bunch

# Different spobs, different biomes
my @kerBiomes = qw (Water Shores Grasslands Highlands Mountains Deserts
		    Badlands Tundra IceCaps);
my @munBiomes = qw (FarsideCrater HighlandCraters Highlands MidlandCraters
		    Midlands NorthernBasin NorthwestCrater PolarCrater
		    PolarLowlands Poles SouthwestCrater TwinCraters Canyons
		    EastCrater EastFarsideCrater);
my @minBiomes = qw (Flats GreatFlats GreaterFlats Highlands LesserFlats
		    Lowlands Midlands Poles Slopes);
my @mohBiomes = qw (NorthPole NorthernSinkholeRidge NorthernSinkhole Highlands
		    Midlands MinorCraters CentralLowlands WesternLowlands
		    SouthWesternLowlands SouthEasternLowlands Canyon
		    SouthPole);
my @eveBiomes = qw (Poles ExplodiumSea Lowlands Midlands Highlands Peaks
		    ImpactEjecta);
my @gilBiomes = qw (Lowlands Midlands Highlands);
my @dunBiomes = qw (Poles Highlands Midlands Lowlands Craters);
my @ikeBiomes = qw (PolarLowlands Midlands Lowlands EasternMountainRidge
		    WesternMountainRidge CentralMountainRidge
		    SouthEasternMountainRange SouthPole);
my @dreBiomes = qw (Poles Highlands Midlands Lowlands Ridges ImpactEjecta
		    ImpactCraters Canyons);
my @layBiomes = qw (Poles Shores Dunes TheSagenSea);
my @valBiomes = qw (Poles Highlands Midlands Lowlands);
my @tylBiomes = qw (Highlands Midlands Lowlands Mara MajorCrater);
my @bopBiomes = qw (Poles Slopes Peaks Valley Rodges);
my @polBiomes = qw (Poles Lowlands Midlands Highlands);
my @eelBiomes = qw (Poles Glaciers Midlands Lowlands IceCanyons Highlands
		    Craters);


# Am I in a science or recovery loop, or did I leave a recovery loop?
my $ticker = '0';
my $recoTicker = '0';
my $scanTicker = '0';
my $eolTicker = '0';

# Sometimes I use one versus the other, mainly for spacing in averages table
my $recov = 'Recov';
my $recovery = 'recovery';
my $scansat = 'SCANsat';


### Begin!
# Construct sbv hash
while (<DATA>) {
  chomp;
  my @sbvs = split;
  $sbvData{$sbvs[0].$sbvs[1]} = $sbvs[2];
}

# Read in science defs to build prebuild datamatrix for each experiment
open my $defs, '<', "$scidef" or die $!;
while (<$defs>) {
  chomp;

  # Find all the science loops
  if (m/^EXPERIMENT_DEFINITION/) {
    $ticker = 1;
    next;
  }

  # Note when we close out of a loop, nothing valuable after that
  elsif (m/RESULTS/) {
    $ticker = 0;
    next;
  }

  # Skip the first line, remove leading tabs, and assign arrays
  elsif ($ticker == 1) {
    next if m/^\{|^\s+$/;	# Take into account blank lines
    s/^\t//i;

    my ($tmp1,$tmp2) = split /=/;
    $tmp1 =~ s/\s+//g;		# Clean spaces
    $tmp2 =~ s/\s+//g;		# Also fix default spacing in ScienceDefs.cfg

    if ($tmp1 eq 'id') {
      @testdef = (@testdef,$tmp2);
    } elsif ($tmp1 eq 'situationMask') {
      $tmp2 = binary($tmp2);
      @sitmask = (@sitmask,$tmp2);
    } elsif ($tmp1 eq 'biomeMask') {
      $tmp2 = binary($tmp2);
      @biomask = (@biomask,$tmp2);
    } elsif ($tmp1 eq 'requireAtmosphere') {
      @atmo = (@atmo,'1') if $tmp2 eq 'True'; # Waiting for fix to sciencedefs
      @atmo = (@atmo,'0') if $tmp2 eq 'False';
    } elsif ($tmp1 eq 'dataScale') {
      @dataScale = (@dataScale,$tmp2);
    } elsif ($tmp1 eq 'scienceCap') {
      @scienceCap = (@scienceCap,$tmp2);
    }
  }
}
close $defs or die $!;


# Iterate and decide on conditions, build matrix, gogogo
foreach my $i (0..scalar @testdef - 1) {
  # Array of binary values, only need to do once per test
  my @sits = split //,$sitmask[$i];
  my @bins = split //,$biomask[$i];

  foreach my $planet (0..$planetCount) {

    # Build list of potential situations
    my @situations = qw (Landed Splashed FlyingLow
			 FlyingHigh InSpaceLow InSpaceHigh);

    # Build an array of arrays, nullify alongside @situations
    # Don't forget the KSC/Runway/Launchpad/etc. biomes, but only for landed?
    # Have to somehow deal with eva report while flying over, goo?
    # ;;;;;; ##### FIXME TODO
    # Only three spobs have biomes as of 0.25
    my @biomes = arrayBuild ($planets[$planet]);

    for (my $var = scalar @sits - 1;$var>=0;$var--) {
      my $vara = abs $var-5;
      if ($sits[$vara] == 0) {
	splice @situations, $var, 1;
	splice @biomes, $var, 1;
      } elsif ($bins[$vara] == 0) {
	$biomes[$var] = [ qw (Global)];
      }
    }

    foreach my $sit (0..scalar @situations - 1) {
      # No surface
      next if (($situations[$sit] eq 'Landed') && ($planets[$planet] =~ m/^Kerbol$|^Jool$/));
      # Water
      next if (($situations[$sit] eq 'Splashed') && ($planets[$planet] !~ m/^Kerbin$|^Eve$|^Laythe$/));
      # Atmosphere
      if ($planets[$planet] !~ m/^Kerbin|^Eve|^Duna|^Jool|^Laythe/) {
	next if $situations[$sit] =~ m/^FlyingLow$|^FlyingHigh$/;
	next if $atmo[$i] == 1;
      }

      foreach my $bin (0..scalar @{$biomes[$sit]} - 1) {
	# Use specific data (test, spob, sit, biome) as key to allow specific
	# references and unique overwriting
	my $sbVal = $sbvData{$planets[$planet].$situations[$sit]};
	my $cleft = $sbVal*$scienceCap[$i];
	$dataMatrix{$testdef[$i].$planets[$planet].$situations[$sit].$biomes[$sit][$bin]} = [$testdef[$i],$planets[$planet],$situations[$sit],$biomes[$sit][$bin],$dataScale[$i],'1',$sbVal,'0',$cleft,$cleft,'0'];
      }
    }
  }
}


# Build recovery hash
foreach my $planet (0..$planetCount) {
  my @situations = qw (FlewBy SubOrbited Orbited Surfaced);

  # Kerbin is special of course
  if ($planets[$planet] eq 'Kerbin') {
    $situations[0] = 'Flew';
    pop @situations;
  }

  foreach my $sit (0..scalar @situations - 1) {
    # No surface
    next if (($situations[$sit] eq 'Surfaced') && ($planets[$planet] =~ m/^Kerbol|^Jool/));
    my $sbVal = $sbvData{$planets[$planet].'Recovery'};
    my $cleft = $sbVal*$recoCaps{$situations[$sit]};
    $reco{$planets[$planet].$situations[$sit]} = [$recovery,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
  }
}

# Build SCANsat hash
foreach my $planet (0..$planetCount) {
  my @situations = qw (AltimetryLoRes BiomeAnomaly AltimetryHiRes);

  foreach my $sit (0..scalar @situations - 1) {
    # No surface?  Do scanning
    next if ($planets[$planet] =~ m/^Kerbol|^Jool/);

    # SCANsat sbv values correspond to InSpaceHigh values
    # NOPE!!!  This apparently changed in a recent update to SCANsat, so now
    # it's somewhat less logical.  This will suffice for now
    # FIXME TODO
    my $sbVal = $sbvData{$planets[$planet].'InSpaceHigh'};
    my $cleft = $sbVal*$scanCaps{$situations[$sit]};
    $scan{$planets[$planet].$situations[$sit]} = [$scansat,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
  }
}

open my $file, '<', "$pers" or die $!;
while (<$file>) {
  chomp;

  # Find all the science loops
  if (m/^\t\tScience$/) {
    $ticker = 1;
    next;
  }

  # Note when we close out of a loop
  elsif (m/\t\t\}/) {
    $ticker = 0;
    next;
  }

  # Skip the first line, remove leading tabs, and assign arrays
  elsif ($ticker == 1) {
    next if m/^\t\t\{/;
    s/^\t\t\t//i;
    my ($tmp1,$tmp2) = split /=/;
    $tmp1 =~ s/\s+//g;		# Clean spaces
    $tmp2 =~ s/\s+//g;
    $tmp2 =~ s/Sun/Kerbol/g;

    if ($tmp1 eq 'id') {
      $eolTicker = 0;

      # Replace recovery and SCANsat data here, why not?
      if ($tmp2 =~ m/^$recovery/) {
	$recoTicker = 1;
	$tmp2 =~ s/(Flew[By]?|SubOrbited|Orbited|Surfaced)/\@$1/g;
	@pieces = (split /@/, $tmp2);
      } elsif ($tmp2 =~ m/^$scansat/) {
	$scanTicker = 1;
	$tmp2 =~ s/InSpaceHighsurface$//g;
	$tmp2 =~ s/^$scansat(.*)\@(.*)/$scansat\@$2\@$1/g;
	@pieces = (split /@/, $tmp2);
      } else {
	($recoTicker,$scanTicker) = (0,0);
	# Watch out for srf landed/splashed, InSpaceHigh/Low, FlyingHigh/Low
	$tmp2 =~ s/Srf(Landed|Splashed)/\@$1\@/g;
	$tmp2 =~ s/(InSpace|Flying)(Low|High)/\@$1$2\@/g;
	@pieces = (split /@/, $tmp2);
      }

      # Ensure arrays are the same length
      push @test, $pieces[0];
      push @spob, $pieces[1];
      push @where, $pieces[2];
      push @biome, $pieces[3] // 'Global'; # global biomes
    } elsif ($tmp1 =~ m/^title/) {
      @title = (@title,$tmp2);
    } elsif ($tmp1 =~ m/^dsc/) {
      @dsc = (@dsc,$tmp2);
    } elsif ($tmp1 =~ m/^scv/) {
      @scv = (@scv,$tmp2);
    } elsif ($tmp1 =~ m/^sbv/) {
      @sbv = (@sbv,$tmp2);
    } elsif ($tmp1 =~ m/^sci/) {
      @sci = (@sci,$tmp2);
    } elsif ($tmp1 =~ m/^cap/) {
      @cap = (@cap,$tmp2);
      $eolTicker = 1;
    }

    # Build hash holding recovery for SCANsat data
    if (($recoTicker == 1) && ($eolTicker == 1)) {
      my $cleft = sprintf '%.2f', 100*$sci[-1]/$cap[-1];
      $reco{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$cleft];
    } elsif (($scanTicker == 1) && ($eolTicker == 1)) {
      my $cleft = sprintf '%.2f', 100*$sci[-1]/$cap[-1];
      $scan{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$cleft];
    }

    # Not sure what do?  ;;;;;; ##### FIXME TODO
    next;
  }
}
close $file or die $!;

# Build the matrix
foreach (0..scalar @test - 1) {
  next if $test[$_] =~ m/^SCANsat|^asteroid/;
  if ($biome[$_]) {
    if (($test[$_] !~ m/$recovery/i) && ($biome[$_] !~ m/^KSC|^Runway|^LaunchPad|^VAB/)) {
      my $cleft = sprintf '%.2f', 100*$sci[$_]/$cap[$_];
      $dataMatrix{$test[$_].$spob[$_].$where[$_].$biome[$_]} = [$test[$_],$spob[$_],$where[$_],$biome[$_],$dsc[$_],$scv[$_],$sbv[$_],$sci[$_],$cap[$_],$cap[$_]-$sci[$_],$cleft];
    }
  }
}

###
### Begin the printing process!
###
my @header = qw [Experiment Spob Condition dsc scv sbv sci cap Left Perc.Accom];

## Prepare fancy-schmancy Excel workbook
# Create new workbook
my $workbook = Excel::Writer::XLSX->new( "$outfile" );
#my $workbook = Excel::Writer::XLSX->new( "tmp" );
# Bold for headers, red for science left, green for stupidly small values
my $bold = $workbook->add_format();
my $bgRed = $workbook->add_format();
my $bgGreen = $workbook->add_format();

# Turn off formatting if so desired
if (!$opts{n}) {
  $bold->set_bold();
  $bgRed->set_bg_color( 'red' );
  $bgGreen->set_bg_color( 'green' );
}

# Generate each worksheet with proper header
# Subroutine these ;;;;;; ##### FIXME TODO
$workVars{$recov} = [$workbook->add_worksheet( 'Recovery' ), 1];
$workVars{$recov}[0]->write( 0, 0, \@header, $bold );

$workVars{$scansat} = [$workbook->add_worksheet( 'SCANsat' ), 1];
$workVars{$scansat}[0]->write( 0, 0, \@header, $bold );

$header[1] = 'Condition';
$header[2] = 'Biome';

foreach my $planet (0..$planetCount) {
  # Interpolate via " instead of '
  $workVars{$planets[$planet]} = [$workbook->add_worksheet( "$planets[$planet]" ), 1];
  $workVars{$planets[$planet]}[0]->write( 0, 0, \@header, $bold );
}

# Recovery widths, manually determined
# Subroutine these ;;;;;; ##### FIXME TODO
$workVars{$recov}[0]->set_column( 0, 0, 9.17 );
$workVars{$recov}[0]->set_column( 1, 1, 6.5 );
$workVars{$recov}[0]->set_column( 2, 2, 9 );
# SCANsat widths, manually determined
$workVars{$scansat}[0]->set_column( 0, 0, 9.17 );
$workVars{$scansat}[0]->set_column( 1, 1, 6.5 );
$workVars{$scansat}[0]->set_column( 2, 2, 11.83 );
# Stock science widths, manually determined
foreach my $planet (0..$planetCount) {
  $workVars{$planets[$planet]}[0]->set_column( 0, 0, 15.5 );
  $workVars{$planets[$planet]}[0]->set_column( 1, 1, 9.67 );
  $workVars{$planets[$planet]}[0]->set_column( 2, 2, 8.5 );
}


## Actually print everybody!
open my $csv, '>', "$csvOut" or die $!;
writeToCSV(\@header) if $opts{c};

# Stock science
foreach my $key (sort sitSort keys %dataMatrix) {
  # Splice out planet name so it's not repetitive
  my $planet = splice @{$dataMatrix{$key}}, 1, 1;
  writeToExcel($planet,\@{$dataMatrix{$key}},$key,\%dataMatrix);
  writeToCSV(\@{$dataMatrix{$key}}) if $opts{c};

  if ($opts{t}) {
    buildScienceData($key,$dataMatrix{$key}[0],\%testData,\%dataMatrix);
  } elsif ($opts{a}) {
    buildScienceData($key,$planet,\%spobData,\%dataMatrix);
  }
}
# Recovery
foreach my $key (sort { specialSort($a, $b, \%reco) } keys %reco) {
  writeToExcel($recov,\@{$reco{$key}},$key,\%reco);
  writeToCSV(\@{$reco{$key}}) if $opts{c};

  if ($opts{t}) {
    # Neater spacing in test averages output
    buildScienceData($key,$recovery,\%testData,\%reco);
  } elsif ($opts{a}) {
    buildScienceData($key,$recov,\%spobData,\%reco);
  }
}
# SCANsat
foreach my $key (sort { specialSort($a, $b, \%scan) } keys %scan) {
  writeToExcel($scansat,\@{$scan{$key}},$key,\%scan);
  writeToCSV(\@{$scan{$key}}) if $opts{c};

  if ($opts{t}) {
    # Neater spacing in test averages output
    buildScienceData($key,$scansat,\%testData,\%scan);
  } elsif ($opts{a}) {
    buildScienceData($key,$scansat,\%spobData,\%scan);
  }
}
close $csv or die $!;


## Sorting of different average tables
# Ensure the -t flag supersedes -a if both are given
if ($opts{a} || $opts{t}) {
  my $string = "Average science left:\n\n";
  my ($tmpHashRef,$tmpArrayRef);

  if ($opts{t}) {
    $string .= "Test\t";
    $tmpHashRef = \%testData;
    $tmpArrayRef = \@testdef if !$opts{s};
  } elsif ($opts{a}) {
    $string .= 'Spob';
    $tmpHashRef = \%spobData;
    $tmpArrayRef = \@planets if !$opts{s};
  }
  $string .= "\tAvg/exp\tTotal\tCompleted\n";
  print "$string";

  if ($opts{p}) {
    average3($tmpHashRef);
  } elsif ($opts{s}) {
    average2($tmpHashRef);
  } else {
    average1($tmpHashRef,$tmpArrayRef);
  }
}


### SUBROUTINES
# Convert string to binary, pad to six digits
sub binary
  {
    my $ones = sprintf '%b',shift;
    while (length($ones)<6) {
      $ones = '0'.$ones;
    }
    return $ones;
  }

# Create array of arrays for spob-specific @biomes
# More efficiently pass first three letters for array assignment
# FIXME TODO
sub arrayBuild
  {
    my $plane = shift;
    my @tmpArray;
    if ($plane eq 'Kerbin') {
      @tmpArray = ([@kerBiomes])x6;
    } elsif ($plane eq 'Mun') {
      @tmpArray = ([@munBiomes])x6;
    } elsif ($plane eq 'Minmus') {
      @tmpArray = ([@minBiomes])x6;
    } else {
      @tmpArray = ([qw (Global)])x6;
    }

    return @tmpArray;
  }

## Custom sort order, adapted from:
## http://stackoverflow.com/a/8171591/2521092
# Kerbin and moons come first, then Kerbol, then proper sorting of conditions,
# matches worksheets
# Incorporate KSC FIXME TODO
sub specialSort
  {
    my ($a,$b,$specRef) = @_;
    my @input = ($a, $b);	# Keep 'em separate, avoid expr version of map

    my @specOrder = @planets;
    my %spec_order_map = map { $specOrder[$_] => $_ } 0 .. $#specOrder;
    my $sord = join q{|}, @specOrder;
    my @condOrder = qw (Flew SubOrbited Orbited Surfaced AltimetryLoRes AltimetryHiRes BiomeAnomaly);
    my %cond_order_map = map { $condOrder[$_] => $_ } 0 .. $#condOrder;
    my $cord = join q{|}, @condOrder;

    my ($x,$y) = map {/^($sord)/} @input;
    my ($v,$w) = map {/($cord)/} @input;

    if ($opts{p}) {
      ${$specRef}{$b}[9] <=> ${$specRef}{$a}[9] || $a cmp $b || $cond_order_map{$v} <=> $cond_order_map{$w};
    } elsif ($opts{s}) {
      ${$specRef}{$b}[8] <=> ${$specRef}{$a}[8] || $a cmp $b || $cond_order_map{$v} <=> $cond_order_map{$w};
    } else {
      $spec_order_map{$x} <=> $spec_order_map{$y} || $cond_order_map{$v} <=> $cond_order_map{$w};
    }
  }

# Sort alphabetically by test, then specifically by situation, then
# alphabetically by biome
sub sitSort
  {
    my @input = ($a, $b);	# Keep 'em separate, avoid expr version of map

    my @sitOrder = qw (Landed Splashed FlyingLow FlyingHigh InSpaceLow InSpaceHigh);
    my %sit_order_map = map { $sitOrder[$_] => $_ } 0..$#sitOrder;
    my $sord = join q{|}, @sitOrder;

    # Test
    my ($v,$w) = ($a,$b);
    $v =~ s/^(.*)(Landed|Splashed|FlyingLow|FlyingHigh|InSpaceLow|InSpaceHigh).*/$1/i;
    $w =~ s/^(.*)(Landed|Splashed|FlyingLow|FlyingHigh|InSpaceLow|InSpaceHigh).*/$1/i;
    # Biome
    my ($t,$u) = ($a,$b);
    $t =~ s/^.*(Landed|Splashed|FlyingLow|FlyingHigh|InSpaceLow|InSpaceHigh)(.*)/$2/i;
    $u =~ s/^.*(Landed|Splashed|FlyingLow|FlyingHigh|InSpaceLow|InSpaceHigh)(.*)/$2/i;

    my ($x,$y) = map {/($sord)/} @input;

    if ($opts{p}) {
      $dataMatrix{$b}[10] <=> $dataMatrix{$a}[10] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    } elsif ($opts{s}) {
      $dataMatrix{$b}[9] <=> $dataMatrix{$a}[9] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    } else {
      $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    }
  }

sub writeToCSV
  {
    my $lineRef = shift;

    print $csv join q{,} , @{$lineRef};
    print $csv "\n";
  }

sub writeToExcel
  {
    my ($sheetName,$rowRef,$matrixKey,$hashRef) = @_;

    $workVars{$sheetName}[0]->write_row( $workVars{$sheetName}[1], 0, $rowRef );
    $workVars{$sheetName}[0]->write( $workVars{$sheetName}[1], 8, ${$hashRef}{$matrixKey}[8], $bgRed ) if ${$hashRef}{$matrixKey}[8] > 0;
    $workVars{$sheetName}[0]->write( $workVars{$sheetName}[1], 4, ${$hashRef}{$matrixKey}[4], $bgGreen ) if ((${$hashRef}{$matrixKey}[4] < 0.001) && (${$hashRef}{$matrixKey}[4] >0));
    $workVars{$sheetName}[1]++;
    return;
  }

# Build data hashes for averages
sub buildScienceData
  {
    my ($key,$ind,$dataRef,$hashRef) = @_;

    ${$dataRef}{$ind}[0] += ${$hashRef}{$key}[8];
    ${$dataRef}{$ind}[1]++;
    ${$dataRef}{$ind}[2] += ${$hashRef}{$key}[7];

    return;
  }


# Alphabeticalish averages
sub average1
  {
    my $hashRef = shift;
    my $arrayRef = shift;

    if ($opts{t}) {
      push @{$arrayRef}, $recovery; # Neater spacing in test averages output
      push @{$arrayRef}, $scansat;  # Neater spacing in test averages output
      @{$arrayRef} = sort @{$arrayRef};
    }

    foreach my $index (0..scalar @{$arrayRef} - 1) {
      printAverageTable(${$arrayRef}[$index],$hashRef);
    }

    if (!$opts{t}) {
      printAverageTable($recov,$hashRef);
      printAverageTable($scansat,$hashRef);
    }

    return;
  }

# Averages sorted by total remaining science
sub average2
  {
    my $hashRef = shift;

    foreach my $key (sort {${$hashRef}{$b}[0] <=> ${$hashRef}{$a}[0] || $a cmp $b} keys %{$hashRef}) {
      printAverageTable($key,$hashRef);
    }

    return;
  }

# Averages sorted by percent accomplished
sub average3
  {
    my $hashRef = shift;

    foreach my $key (sort {((${$hashRef}{$b}[2]-${$hashRef}{$b}[0])/${$hashRef}{$b}[2]) <=> ((${$hashRef}{$a}[2]-${$hashRef}{$a}[0])/${$hashRef}{$a}[2]) || $a cmp $b} keys %{$hashRef}) {
      printAverageTable($key,$hashRef);
    }

    return;
  }

# Handle printing of the averages table
sub printAverageTable
  {
    my @placeHolder = @_;
    my $ind = $placeHolder[0];
    my %hash = %{$placeHolder[1]};

    my $indL = substr $ind, 0, 14; # Neater spacing in test averages output
    my $avg = $hash{$ind}[0]/($hash{$ind}[1]);
    my $remains = $hash{$ind}[2] - $hash{$ind}[0];
    my $per = 100*$remains/$hash{$ind}[2];

    printf "%s\t%.0f\t%.0f\t%.0f\n", $indL, $avg, $hash{$ind}[0], $per;

    return;
  }

#### Usage statement ####
# Use POD or whatever?
# Escapes not necessary but ensure pretty colors
# Final line must be unindented?
sub usage
  {
    print <<USAGE;
Usage: $0 [-atsnchH -u <savefile_name>]
      -a Display average science left for each planet.
      -t Display average science left for each experiment type.
      -s Sort output by science left, including averages from -a and -t flags.
      -p Sort output by percent science accomplished, including averages from
         the -a and -t flags.  Supersedes the -s flag.
      -n Turn off formatted printing (i.e., colors and bolding).
      -c Output data to csv file as well
      -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -h or H Print this message.
USAGE
    return;
  }


## The lines below do not represent Perl code, and are not examined by the
## compiler.  Rather, they are the default celestial body multipliers,
## represented as sbv in persistent.sfs (and presumably defined *somewhere*).
## These values decay along with experiment progress to determine science
## output.  In reality there is only one value each for atmo and space, but it
## was easier to just duplicate them here rather than build another loop.
__END__
Kerbol InSpaceLow 11
  Kerbol InSpaceHigh 11
  Kerbol Recovery 4
  Kerbin Landed 0.3
  Kerbin Splashed 0.4
  Kerbin FlyingLow 0.7
  Kerbin FlyingHigh 0.7
  Kerbin InSpaceLow 1
  Kerbin InSpaceHigh 1
  Kerbin Recovery 1
  Mun Landed 4
  Mun InSpaceLow 3
  Mun InSpaceHigh 3
  Mun Recovery 2
  Minmus Landed 5
  Minmus InSpaceLow 4
  Minmus InSpaceHigh 4
  Minmus Recovery 2.5
  Moho Landed 10
  Moho InSpaceLow 8
  Moho InSpaceHigh 8
  Moho Recovery 7
  Eve Landed 8
  Eve Splashed 8
  Eve FlyingLow 6
  Eve FlyingHigh 6
  Eve InSpaceLow 7
  Eve InSpaceHigh 7
  Eve Recovery 5
  Gilly Landed 9
  Gilly InSpaceLow 8
  Gilly InSpaceHigh 8
  Gilly Recovery 6
  Duna Landed 8
  Duna FlyingLow 5
  Duna FlyingHigh 5
  Duna InSpaceLow 7
  Duna InSpaceHigh 7
  Duna Recovery 5
  Ike Landed 8
  Ike InSpaceLow 7
  Ike InSpaceHigh 7
  Ike Recovery 5
  Dres Landed 8
  Dres InSpaceLow 7
  Dres InSpaceHigh 7
  Dres Recovery 6
  Jool FlyingLow 12
  Jool FlyingHigh 12
  Jool InSpaceLow 7
  Jool InSpaceHigh 7
  Jool Recovery 6
  Laythe Landed 14
  Laythe Splashed 12
  Laythe FlyingLow 11
  Laythe FlyingHigh 11
  Laythe InSpaceLow 9
  Laythe InSpaceHigh 9
  Laythe Recovery 8
  Vall Landed 12
  Vall InSpaceLow 9
  Vall InSpaceHigh 9
  Vall Recovery 8
  Tylo Landed 12
  Tylo InSpaceLow 10
  Tylo InSpaceHigh 10
  Tylo Recovery 8
  Bop Landed 12
  Bop InSpaceLow 9
  Bop InSpaceHigh 9
  Bop Recovery 8
  Pol Landed 12
  Pol InSpaceLow 9
  Pol InSpaceHigh 9
  Pol Recovery 8
  Eeloo Landed 15
  Eeloo InSpaceLow 12
  Eeloo InSpaceHigh 12
  Eeloo Recovery 10
