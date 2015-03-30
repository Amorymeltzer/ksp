#!/usr/bin/env perl
# parseScience.pl by Amory Meltzer
# v0.93.1
# https://github.com/Amorymeltzer/ksp
# Parse a KSP persistent.sfs file, report science information
# Sun represented as Kerbol
# Leftover science in red, candidates for manual cleanup in green

## Ignores KSC/LaunchPad/Runway/etc. "biomes", asteroids
## Can you do srfsplashed in every biome on other planets with water?

### Add support for:
## KSC biomes
## Asteroids

### FIXES, TODOS
## SCANsat allows Sun scanning?!
## Use Cwd even necessary for config processing?!
## Option to skip dsc, sbv, etc. stuff?
## Windows path to Gamedata/pers/scidefs/etc.?
## Option to pull KSC stuff in/out of Kerbin?
## Option to print averages table to file
## User-specified KSP location?
## Incorporate InSpaceLow/High, etc. cutoffs somehow
### Print into excel?  New header?

## Turn cascading tmp1/2 elsifs into hash lookup?  Might revert above
## Cleanup data/test hashes, the order of the data is unintuitive
## Cleanup var/vara/etc. crap.  Better commenting.

## Biomes are hardcoded, would be nice to pull from somewhere
## Same for sbv, multipliers, etc.


use strict;
use warnings;
use diagnostics;

use Getopt::Std;
use Cwd;
use FindBin;
use Excel::Writer::XLSX;

# Parse command line options
my %opts = ();
getopts('aAtTsSpPnNcCiIkKu:Uf:h', \%opts);

if ($opts{h}) {
  usage();
  exit;
}

### ENVIRONMENT VARIABLES
my $dotfile;			# Preference file
# Replaced from the %opts table
my %opt = (
	   username => 0,
	   average => 0,
	   tests => 0,
	   scienceleft => 0,
	   percentdone => 0,
	   noformat => 0,
	   csv => 0,
	   includeSCANsat => 0,
	   ksckerbin => 0
	  );

## .parsesciencerc config file
# Useful shorthands for finding files
my $rc = 'parsesciencerc';	# Config dotfile name of choice.  Wordy.
my $cwd = cwd();		# Current working directory
my $scriptDir = $FindBin::Bin;	# Directory of this script
my $home = $ENV{HOME};		# MAGIC hash with user env variables for $home

# Round up the usual suspects, all superseded by commandline flag
my @dotLocales = ("$cwd/.$rc","$scriptDir/.$rc","$home/.$rc","$home/.config/parseScience/$rc");
if ($opts{f} && -e $opts{f}) {
  $dotfile = $opts{f};
} else {
  foreach my $place (@dotLocales) {
    if (-e $place) {
      $dotfile = $place;
      last;
    }
  }
}

# Parse config file
if ($dotfile) {
  open my $dot, '<', "$dotfile" or die $!;
  while (<$dot>) {
    chomp;

    next if m/^#/g;		# Ignore comments
    next if !$_;		# Ignore blank lines

    if (!m/^.+ = .+/) {		# Ignore and warn on malformed entries
      warnNicely("Malformed entry '$_' at line $. of $dotfile.  Skipping...");
      next;
    }

    s/ //g;
    my @config = split /=/;

    if ($config[0] eq 'username') {
      $opt{$config[0]} = $config[1];
    } elsif ($config[1] eq 'true') {
      $opt{$config[0]} = 1;
    } elsif ($config[1] eq 'false') {
      $opt{$config[0]} = 0;
    } else {
      warnNicely("Unknown option '$config[0]' at line $. of $dotfile.  Skipping...");
      next;
    }
  }
  close $dot or die $!;
}


### FILE DEFINITIONS
# Change this to match the location of your KSP install
my $path = '/Applications/KSP_osx';
my $scidef = 'ScienceDefs.cfg';
my $pers = 'persistent.sfs';

if ($opts{u}) {
  $opt{username} = $opts{u};
}

# Overwrite config file options if the corresponding flag is on the commandline
# Negated options always take precedence
my @negatableOpts = keys %opt;
foreach my $negate (@negatableOpts) {
  my $ng8 = substr $negate, 0, 1; # Short key for %opts
  my $Ung8 = uc $ng8;		  # Uppercase key for negated option

  if ($opts{$Ung8}) {
    $opt{$negate} = 0;
  } elsif ($opts{$ng8}) {
    $opt{$negate} = $opts{$ng8};
  }
}

if ($opt{username}) {
  $scidef = "$path/GameData/Squad/Resources/ScienceDefs.cfg";
  $pers = "$path/saves/$opt{username}/persistent.sfs";
}

# Test files for existance
warnNicely("No ScienceDefs.cfg file found at $scidef", 1) if !-e $scidef;
warnNicely("No persistent.sfs file found at $pers\n", 1) if !-e $pers;

my $outfile = 'scienceToDo.xlsx';
my $csvFile = 'scienceToDo.csv';


### GLOBAL VARIABLES
my %dataMatrix;		      # Hold stock data
my %reco;		      # Separate hash for craft recovery
my %scan;		      # Separate hash for SCANsat
my %sbvData;		      # Hold sbv values from END data
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

my @planets = qw (Kerbin KSC Mun Minmus Kerbol Moho Eve Gilly Duna Ike Dres
		  Jool Laythe Vall Tylo Bop Pol Eeloo);
my $planetCount = scalar @planets - 1; # Use this a bunch

# Different spobs, different biomes
my %universe = (
		Kerbin => [ qw (Water Shores Grasslands Highlands Mountains
				Deserts Badlands Tundra IceCaps) ],
		KSC => [ qw (KSC Administration AstronautComplex Crawlerway
			     FlagPole LaundPad MissionControl R&D
			     R&DCentralBuilding R&DCornerLab R&DMainBuilding
			     R&DOBservatory R&DSideLab R&DSmallLab R&DTanks
			     R&DWindTunnel Runway SPH SPHMainBuilding
			     SPHRoundTank SPHTanks SPHWaterTower
			     TrackingStation TrackingStationDishEast
			     TrackingStationDishNorth TrackingStationSouth
			     TrackingStationHub VAB VABMainBuilding
			     VABPodMemorial VABRoundTank VABTanks) ],
		Mun => [ qw (FarsideCrater HighlandCraters Highlands
			     MidlandCraters Midlands NorthernBasin
			     NorthwestCrater PolarCrater PolarLowlands Poles
			     SouthwestCrater TwinCraters Canyons EastCrater
			     EastFarsideCrater) ],
		Minmus => [ qw (Flats GreatFlats GreaterFlats Highlands
				LesserFlats Lowlands Midlands Poles Slopes) ],
		Kerbol => [ qw (Global) ],
		Moho => [ qw (NorthPole NorthernSinkholeRidge NorthernSinkhole
			      Highlands Midlands MinorCraters CentralLowlands
			      WesternLowlands SouthWesternLowlands
			      SouthEasternLowlands Canyon SouthPole) ],
		Eve => [ qw (Poles ExplodiumSea Lowlands Midlands Highlands
			     Peaks ImpactEjecta) ],
		Gilly => [ qw (Lowlands Midlands Highlands) ],
		Duna => [ qw (Poles Highlands Midlands Lowlands Craters) ],
		Ike => [ qw (PolarLowlands Midlands Lowlands
			     EasternMountainRidge WesternMountainRidge
			     CentralMountainRidge SouthEasternMountainRange
			     SouthPole) ],
		Dres => [ qw (Poles Highlands Midlands Lowlands Ridges
			      ImpactEjecta ImpactCraters Canyons) ],
		Jool => [ qw (Global) ],
		Laythe => [ qw (Poles Shores Dunes TheSagenSea) ],
		Vall => [ qw (Poles Highlands Midlands Lowlands) ],
		Tylo => [ qw (Highlands Midlands Lowlands Mara MajorCrater) ],
		Bop => [ qw (Poles Slopes Peaks Valley Rodges) ],
		Pol => [ qw (Poles Lowlands Midlands Highlands) ],
		Eeloo => [ qw (Poles Glaciers Midlands Lowlands IceCanyons
			       Highlands Craters) ]
	       );

# Various situations you may find yourself in
my @stockSits = qw (Landed Splashed FlyingLow FlyingHigh InSpaceLow InSpaceHigh);
my @recoSits = qw (Flew FlewBy SubOrbited Orbited Surfaced);
my @scanSits = qw (AltimetryLoRes AltimetryHiRes BiomeAnomaly);

# Reverse-engineered caps for recovery missions and SCANsat data.  SubOrbited and
# Orbited are messed up - the default values from Kerbin are inverted
# elsewhere.  All SCANsat caps are 20
my %recoCap = (
	       Flew => 6,
	       FlewBy => 7.2,
	       SubOrbited => 9.6,
	       Orbited => 12,
	       Surfaced => 18
	      );
my $scanCap = 20;

# Am I in a science, recovery, or SCANsat loop?
my $ticker = '0';
my $recoTicker = '0';
my $scanTicker = '0';

# Sometimes I use one versus the other, mainly for spacing in averages table
my $recov = 'Recov';
my $recovery = 'recovery';
my $scansat = 'SCANsat';
my $scansatMap = 'SCANsatMapping';


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


## Iterate and decide on conditions, build matrix, gogogo
# Build stock science hash
foreach my $i (0..scalar @testdef - 1) {
  # Array of binary values, only need to do once per test
  my @sits = split //,$sitmask[$i];
  my @bins = split //,$biomask[$i];

  foreach my $planet (0..$planetCount) {

    # Build list of potential situations
    my @situations = @stockSits;

    # Create array of arrays for spob-specific biomes, nullify w/ @situations
    my @biomes = ([@{$universe{$planets[$planet]}}])x6;

    # KSC biomes are SrfLanded only
    if ($planets[$planet] eq 'KSC') {
      @situations = qw (Landed);
    } else {
      for (my $var = scalar @sits - 1;$var>=0;$var--) {
	my $vara = abs $var-5;
	if ($sits[$vara] == 0) {
	  splice @situations, $var, 1;
	  splice @biomes, $var, 1;
	} elsif ($bins[$vara] == 0) {
	  $biomes[$var] = [ qw (Global)];
	}
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

      # Inconvenient, ruined by the cleaner funciton later
      if ($planets[$planet] eq 'KSC' && $opt{ksckerbin}) {
	$planets[$planet] = 'Kerbin';
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

# UGLY as sin ;;;;;; ##### FIXME TODO
if ($opt{ksckerbin}) {
  splice @planets, 1, 1;
  $planetCount--;
}

# Build recovery hash
foreach my $planet (0..$planetCount) {
  my @situations = @recoSits;

  # Only one of Flew or FlewBy
  shift @situations;
  # Kerbin is special of course
  if ($planets[$planet] eq 'Kerbin') {
    $situations[0] = 'Flew';
    pop @situations;
  }

  foreach my $sit (0..scalar @situations - 1) {
    next if $planets[$planet] eq 'KSC';
    # No surface
    next if (($situations[$sit] eq 'Surfaced') && ($planets[$planet] =~ m/^Kerbol|^Jool/));
    my $sbVal = $sbvData{$planets[$planet].'Recovery'};
    my $cleft = $sbVal*$recoCap{$situations[$sit]};
    $reco{$planets[$planet].$situations[$sit]} = [$recovery,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
  }
}

# Build SCANsat hash
if ($opt{includeSCANsat}) {
  foreach my $planet (0..$planetCount) {
    my @situations = @scanSits;

    foreach my $sit (0..scalar @situations - 1) {
      # No surface?  Do scanning
      next if ($planets[$planet] =~ m/^Kerbol|^Jool|^KSC/);

      my $sbVal = $sbvData{$planets[$planet].'InSpaceHigh'};
      my $cleft = $sbVal*$scanCap;
      $scan{$planets[$planet].$situations[$sit]} = [$scansat,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
    }
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
      # Replace recovery and SCANsat data here, why not?
      if ($tmp2 =~ m/^$recovery/) {
	$recoTicker = 1;
	$tmp2 =~ s/(Flew[By]?|SubOrbited|Orbited|Surfaced)/\@$1/g;
	@pieces = (split /@/, $tmp2);
      } elsif ($opt{includeSCANsat} && $tmp2 =~ m/^$scansat/) {
	$scanTicker = 1;
	$tmp2 =~ s/^$scansat(.*)\@(.*)InSpaceHighsurface$/$scansat\@$2\@$1/g;
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

      # Build recovery and SCANsat data hashes
      if ($recoTicker == 1) {
	my $cleft = calcPerc($sci[-1],$cap[-1]);
	$reco{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$cleft];
      } elsif ($opt{includeSCANsat} && $scanTicker == 1) {
	my $cleft = calcPerc($sci[-1],$cap[-1]);
	$scan{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$cleft];
      }
    }

    # Not sure what do?  ;;;;;; ##### FIXME TODO
    next;
  }
}
close $file or die $!;

# Build the matrix
foreach (0..scalar @test - 1) {
  # Exclude tests stored in separate hashes
  next if $test[$_] =~ m/^SCANsat|^asteroid/;
  if ($biome[$_]) {
    if ($test[$_] !~ m/$recovery/i) {
      my $cleft = calcPerc($sci[$_],$cap[$_]);

      # Take KSC out of Kerbin
      if (!$opt{ksckerbin} && $biome[$_] =~ m/^KSC|^Runway|^LaunchPad|^VAB|^SPH|^R&D|^Astronaut|^FlagPole|^Mission|^Tracking|^Crawler|^Administration/) {
	$spob[$_] = 'KSC';
      }

      # Skip over annoying "fake" science expts caused by ScienceAlert
      # For more info see
      # http://forum.kerbalspaceprogram.com/threads/76793-0-90-ScienceAlert-1-8-4-Experiment-availability-feedback-%28December-23%29?p=1671187&viewfull=1#post1671187
      # Might cause problems with KSC biomes later FIXME TODO
      next if !$dataMatrix{$test[$_].$spob[$_].$where[$_].$biome[$_]};

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
if (!$opt{noformat}) {
  $bold->set_bold();
  $bgRed->set_bg_color( 'red' );
  $bgGreen->set_bg_color( 'green' );
}

# Generate each worksheet with proper header
# Subroutine these ;;;;;; ##### FIXME TODO
$workVars{$recov} = [$workbook->add_worksheet( 'Recovery' ), 1];
$workVars{$recov}[0]->write( 0, 0, \@header, $bold );
# Recovery widths, manually determined
# Subroutine these ;;;;;; ##### FIXME TODO
$workVars{$recov}[0]->set_column( 0, 0, 9.17 );
$workVars{$recov}[0]->set_column( 1, 1, 6.5 );
$workVars{$recov}[0]->set_column( 2, 2, 9 );

if ($opt{includeSCANsat}) {
  $workVars{$scansat} = [$workbook->add_worksheet( 'SCANsat' ), 1];
  $workVars{$scansat}[0]->write( 0, 0, \@header, $bold );

  # SCANsat widths, manually determined
  $workVars{$scansat}[0]->set_column( 0, 0, 9.17 );
  $workVars{$scansat}[0]->set_column( 1, 1, 6.5 );
  $workVars{$scansat}[0]->set_column( 2, 2, 11.83 );
}

$header[1] = 'Condition';
$header[2] = 'Biome';

foreach my $planet (0..$planetCount) {
  # Interpolate via " instead of '
  $workVars{$planets[$planet]} = [$workbook->add_worksheet( "$planets[$planet]" ), 1];
  $workVars{$planets[$planet]}[0]->write( 0, 0, \@header, $bold );
}

# Stock science widths, manually determined
foreach my $planet (0..$planetCount) {
  $workVars{$planets[$planet]}[0]->set_column( 0, 0, 15.5 );
  $workVars{$planets[$planet]}[0]->set_column( 1, 1, 9.67 );
  $workVars{$planets[$planet]}[0]->set_column( 2, 2, 8.5 );
}


## Actually print everybody!
open my $csvOut, '>', "$csvFile" or die $! if $opt{csv};
writeToCSV(\@header) if $opt{csv};

# Stock science
foreach my $key (sort sitSort keys %dataMatrix) {
  # Splice out planet name so it's not repetitive
  my $planet = splice @{$dataMatrix{$key}}, 1, 1;
  writeToExcel($planet,\@{$dataMatrix{$key}},$key,\%dataMatrix);

  # Add in spob name to csv, only necessary for stock science
  $dataMatrix{$key}[1] .= "\@$planet";
  writeToCSV(\@{$dataMatrix{$key}}) if $opt{csv};

  if ($opt{tests}) {
    buildScienceData($key,$dataMatrix{$key}[0],\%testData,\%dataMatrix);
  } elsif ($opt{average}) {
    buildScienceData($key,$planet,\%spobData,\%dataMatrix);
  }
}
# Recovery
foreach my $key (sort { specialSort($a, $b, \%reco) } keys %reco) {
  writeToExcel($recov,\@{$reco{$key}},$key,\%reco);
  writeToCSV(\@{$reco{$key}}) if $opt{csv};

  if ($opt{tests}) {
    # Neater spacing in test averages output
    buildScienceData($key,$recovery,\%testData,\%reco);
  } elsif ($opt{average}) {
    buildScienceData($key,$recov,\%spobData,\%reco);
  }
}
# SCANsat
if ($opt{includeSCANsat}) {
  foreach my $key (sort { specialSort($a, $b, \%scan) } keys %scan) {
    writeToExcel($scansat,\@{$scan{$key}},$key,\%scan);
    writeToCSV(\@{$scan{$key}}) if $opt{csv};

    if ($opt{tests}) {
      # Neater spacing in test averages output
      buildScienceData($key,$scansatMap,\%testData,\%scan);
    } elsif ($opt{average}) {
      buildScienceData($key,$scansat,\%spobData,\%scan);
    }
  }
}
close $csvOut or die $! if  $opt{csv};


## Sorting of different average tables
# Ensure the -t flag supersedes -a if both are given
if ($opt{average} || $opt{tests}) {
  my $string = "Average science left:\n\n";
  my ($tmpHashRef,$tmpArrayRef);

  if ($opt{tests}) {
    $string .= "Test\t";
    $tmpHashRef = \%testData;
    $tmpArrayRef = \@testdef if !$opt{scienceleft};
  } elsif ($opt{average}) {
    $string .= 'Spob';
    $tmpHashRef = \%spobData;
    $tmpArrayRef = \@planets if !$opt{scienceleft};
  }
  $string .= "\tAvg/exp\tTotal\tCompleted\n";
  print "$string";

  if ($opt{percentdone}) {
    average3($tmpHashRef);
  } elsif ($opt{scienceleft}) {
    average2($tmpHashRef);
  } else {
    average1($tmpHashRef,$tmpArrayRef);
  }
}


### SUBROUTINES
sub warnNicely
  {
    my ($err,$ilynPayne) = @_;
    if ($ilynPayne) {
      print 'ERROR: ';
    } else {
      print 'Warning: ';
    }
    print "$err\n";
    exit if $ilynPayne;
  }

# Convert string to binary, pad to six digits
sub binary
  {
    my $ones = sprintf '%b',shift;
    while (length($ones)<6) {
      $ones = '0'.$ones;
    }
    return $ones;
  }

sub calcPerc {
  my ($sciC,$capC) = @_;
  return sprintf '%.2f', 100*$sciC/$capC;
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
    my @condOrder = (@recoSits, @scanSits);
    my %cond_order_map = map { $condOrder[$_] => $_ } 0 .. $#condOrder;
    my $cord = join q{|}, @condOrder;

    my ($x,$y) = map {/^($sord)/} @input;
    my ($v,$w) = map {/($cord)/} @input;

    if ($opt{percentdone}) {
      ${$specRef}{$b}[9] <=> ${$specRef}{$a}[9] || $a cmp $b || $cond_order_map{$v} <=> $cond_order_map{$w};
    } elsif ($opt{scienceleft}) {
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

    my @sitOrder = @stockSits;
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

    if ($opt{percentdone}) {
      $dataMatrix{$b}[10] <=> $dataMatrix{$a}[10] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    } elsif ($opt{scienceleft}) {
      $dataMatrix{$b}[9] <=> $dataMatrix{$a}[9] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    } else {
      $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
    }
  }

sub writeToCSV
  {
    my $lineRef = shift;

    print $csvOut join q{,} , @{$lineRef};
    print $csvOut "\n";
    return;
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

    if ($opt{tests}) {
      push @{$arrayRef}, $recovery; # Neater spacing in test averages output
      push @{$arrayRef}, $scansatMap if $opt{includeSCANsat};
      @{$arrayRef} = sort @{$arrayRef};
    }

    foreach my $index (0..scalar @{$arrayRef} - 1) {
      printAverageTable(${$arrayRef}[$index],$hashRef);
    }

    if (!$opt{tests}) {
      printAverageTable($recov,$hashRef);
      printAverageTable($scansat,$hashRef) if $opt{includeSCANsat};
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
Usage: $0 [-aAtTsSnNcCiI -h -f path/to/dotfile -u <savefile_name>]
      -a Display average science left for each planet.
      -A Turn off -a.
      -t Display average science left for each experiment type.  Supersedes
         the -a flag.
      -T Turn off-T.
      -s Sort output by science left, including averages from -a and -t flags.
      -S Turn off -S.
      -p Sort output by percent science accomplished, including averages from
         the -a and -t flags.  Supersedes the -s flag.
      -P Turn off -P.
      -n Turn off formatted printing (i.e., colors and bolding).
      -N Turn off -N.
      -c Output data to csv file as well.
      -C Turn off -c.
      -i Include data from SCANsat.
      -I Turn off -i.
      -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -U Turn off -u.
      -f Specify path to config file.  Supersedes a local .parsesciencerc file.
      -h Print this message.
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
  Kerbol InSpaceHigh 2
  Kerbol Recovery 4
  Kerbin Landed 0.3
  Kerbin Splashed 0.4
  Kerbin FlyingLow 0.7
  Kerbin FlyingHigh 0.9
  Kerbin InSpaceLow 1
  Kerbin InSpaceHigh 1.5
  Kerbin Recovery 1
  KSC Landed 0.3
  Mun Landed 4
  Mun InSpaceLow 3
  Mun InSpaceHigh 2
  Mun Recovery 2
  Minmus Landed 5
  Minmus InSpaceLow 4
  Minmus InSpaceHigh 2.5
  Minmus Recovery 2.5
  Moho Landed 10
  Moho InSpaceLow 8
  Moho InSpaceHigh 7
  Moho Recovery 7
  Eve Landed 8
  Eve Splashed 8
  Eve FlyingLow 6
  Eve FlyingHigh 6
  Eve InSpaceLow 7
  Eve InSpaceHigh 5
  Eve Recovery 5
  Gilly Landed 9
  Gilly InSpaceLow 8
  Gilly InSpaceHigh 6
  Gilly Recovery 6
  Duna Landed 8
  Duna FlyingLow 5
  Duna FlyingHigh 5
  Duna InSpaceLow 7
  Duna InSpaceHigh 5
  Duna Recovery 5
  Ike Landed 8
  Ike InSpaceLow 7
  Ike InSpaceHigh 5
  Ike Recovery 5
  Dres Landed 8
  Dres InSpaceLow 7
  Dres InSpaceHigh 6
  Dres Recovery 6
  Jool FlyingLow 12
  Jool FlyingHigh 9
  Jool InSpaceLow 7
  Jool InSpaceHigh 6
  Jool Recovery 6
  Laythe Landed 14
  Laythe Splashed 12
  Laythe FlyingLow 11
  Laythe FlyingHigh 10
  Laythe InSpaceLow 9
  Laythe InSpaceHigh 8
  Laythe Recovery 8
  Vall Landed 12
  Vall InSpaceLow 9
  Vall InSpaceHigh 8
  Vall Recovery 8
  Tylo Landed 12
  Tylo InSpaceLow 10
  Tylo InSpaceHigh 8
  Tylo Recovery 8
  Bop Landed 12
  Bop InSpaceLow 9
  Bop InSpaceHigh 8
  Bop Recovery 8
  Pol Landed 12
  Pol InSpaceLow 9
  Pol InSpaceHigh 8
  Pol Recovery 8
  Eeloo Landed 15
  Eeloo InSpaceLow 12
  Eeloo InSpaceHigh 10
  Eeloo Recovery 10
