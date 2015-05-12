#!/usr/bin/env perl
# parseScience.pl by Amory Meltzer
# v0.95.3
# https://github.com/Amorymeltzer/ksp
# Parse a KSP persistent.sfs file, report science information
# Sun represented as Kerbol
# Leftover science in red, candidates for manual cleanup in green

### Add support for:
## ISRU scanning?
### 10 science for scan*recovery multiplier?

### FIXES, TODOS
## asteroidSamples landed at VAB, etc?
## Can you do srfsplashed in every biome on other planets with water?
## Check Major Crater triplication
## Incorporate multiplier?  Might look weird...
## Default windows/linux path to Gamedata/pers/scidefs/etc.?
### Steam locations

## Given/while?  Requires 5.10.1 or 5.14 so maybe not ideal
## Cleanup data/test hashes, the order of the data is unintuitive


use strict;
use warnings;
# use diagnostics;

use Getopt::Std;
use Cwd;
use FindBin;
use English qw( -no_match_vars );

# Parse command line options
my %opts = ();
getopts('aAtTsSpPiIkKmMcCnNeEoOg:Gu:Uf:h', \%opts);

if ($opts{h}) {
  usage();
  exit;
}

# Correlate commandline flags with their corresponding config options
my %lookup = (
	      g => 'gamelocation',
	      u => 'username',
	      a => 'average',
	      t => 'tests',
	      s => 'scienceleft',
	      p => 'percentdone',
	      i => 'scansat',
	      k => 'ksckerbin',
	      m => 'moredata',
	      c => 'csv',
	      n => 'noformat',
	      e => 'excludeexcel',
	      o => 'outputavgtable'
	     );

### ENVIRONMENT VARIABLES
## Build %opt hash, using above lookup table.  Replaced from the %opts table.
my %opt;
foreach my $key (keys %lookup) {
  $opt{$lookup{$key}} = 0;
}

## .parsesciencerc config file
my $dotfile;
# Useful shorthands for finding files
my $rc = 'parsesciencerc';	# Config dotfile name of choice.  Wordy.
my $cwd = cwd();		# Current working directory
my $scriptDir = $FindBin::Bin;	# Directory of this script
my $home = $ENV{HOME};		# MAGIC hash with user env variables for $home

# Round up the usual suspects, all superseded by commandline flag
my @dotLocales = ("$cwd/.$rc","$scriptDir/.$rc");
# Windows (XP anyway) complains about $home
@dotLocales = (@dotLocales,"$home/.$rc","$home/.config/parseScience/$rc") if $home;
if ($opts{f} && -e $opts{f}) {
  $dotfile = $opts{f};
} else {
  foreach my $place (@dotLocales) {
    $place =~ s/\//\\/g if $OSNAME eq 'MSWin32';
    if (-e $place) {
      $dotfile = $place;
      last;
    }
  }
}

# Parse config file
if ($dotfile) {
  open my $dot, '<', "$dotfile" or die $ERRNO;
  while (<$dot>) {
    chomp;

    next if m/^#/g;		# Ignore comments
    next if !$_;		# Ignore blank lines

    if (!m/^.+ = .+/) {		# Ignore and warn on malformed entries
      warnNicely("Malformed entry '$_' at line $NR of $dotfile.  Skipping...");
      next;
    }

    s/ //g;
    my @config = split /=/;

    if ($config[0] =~ /username|gamelocation/) {
      $opt{$config[0]} = $config[1];
    } elsif ($config[1] eq 'true') {
      $opt{$config[0]} = 1;
    } elsif ($config[1] eq 'false') {
      $opt{$config[0]} = 0;
    } else {
      warnNicely("Unknown option '$config[0]' at line $NR of $dotfile.  Skipping...");
      next;
    }
  }
  close $dot or warn $ERRNO;
}


# This serves two functions: Not only does it override config valyes with
# commandline flags, but it properly negates any flags anywhere.  Ensures
# negation precedence through the inverted sort
foreach my $flag (sort {$b cmp $a} keys %opts) {
  my $flagUC = uc $flag;
  my $flagLC = lc $flag;
  next if !$lookup{$flagLC};

  if ($flag eq $flagUC) {
    $opt{$lookup{$flagLC}} = 0;
  } elsif ($opts{$flag}) {
    $opt{$lookup{$flagLC}} = $opts{$flag};
  }
}


# Don't bother outputting average file if there ain't any averages to save
if (!$opt{average} && !$opt{tests}) {
  $opt{outputavgtable} = 0;
  warnNicely('outputavgtable option given but no data table selected (-a or -t).  Skipping...');
}


### FILE DEFINITIONS
my $scidef;
my $pers;
my $path;
my $scidefName = 'ScienceDefs.cfg';
my $persName = 'persistent.sfs';
my $gdsr = 'GameData/Squad/Resources/';

# Build and iterate through all potential options
my @scidefLocales = ("$cwd/$scidefName","$scriptDir/$scidefName");
my @persLocales = ("$cwd/$persName","$scriptDir/$persName");

if (!$opt{gamelocation}) {
  if ($OSNAME eq 'darwin') {
    $path = '/Applications/KSP_osx/';
  } elsif ($OSNAME eq 'linux') {
    $path = '/Applications/KSP_linux/';
  } elsif ($OSNAME eq 'MSWin32') {
    $path = 'C:/Program Files/KSP-win/';
  }
} else {
  $path = $opt{gamelocation};
  @scidefLocales = ($path.$gdsr.$scidefName,@scidefLocales);
}

if ($opt{username}) {
  @scidefLocales = ($path.$gdsr.$scidefName,@scidefLocales);
  @persLocales = ($path."saves/$opt{username}/".$persName,@persLocales);
}

# Test files for existance
# Should probably subroutine this FIXME TODO
# ScienceDefs.cfg
my $last_scidef = pop @scidefLocales;
foreach my $place (@scidefLocales) {
  $scidef = $place;
  if (-e $scidef) {
    last;
  } else {
    warnNicely("No ScienceDefs.cfg file found at $scidef");
  }
}
if (!-e $scidef) {
  $scidef = $last_scidef;
  warnNicely("No ScienceDefs.cfg file found at $scidef",1) if !-e $scidef;
}
# persistent.sfs
my $last_pers = pop @persLocales;
foreach my $place (@persLocales) {
  $pers = $place;
  if (-e $pers) {
    last;
  } else {
    warnNicely("No persistent.sfs file found at $pers\n");
  }
}
if (!-e $pers) {
  $pers = $last_pers;
  warnNicely("No persistent.sfs file found at $pers",1) if !-e $pers;
}

if (!$opt{excludeexcel}) {
  require Excel::Writer::XLSX;
}

my $outfile = 'scienceToDo.xlsx';
my $csvFile = 'scienceToDo.csv';
my $avgFile = 'average_table.txt';


### GLOBAL VARIABLES
my %dataMatrix;		      # Stock data
my %reco;		      # Craft recovery data
my %scan;		      # SCANsat data
my %sbvData;		      # sbv values from END data
my %workVars;		      # Hash of arrays to hold worksheets, current row
my %spobData;		      # Science per spob
my %testData;		      # Science per test

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
			     FlagPole LaunchPad MissionControl R&D
			     R&DCentralBuilding R&DCornerLab R&DMainBuilding
			     R&DObservatory R&DSideLab R&DSmallLab R&DTanks
			     R&DWindTunnel Runway SPH SPHMainBuilding
			     SPHRoundTank SPHTanks SPHWaterTower
			     TrackingStation TrackingStationDishEast
			     TrackingStationDishNorth TrackingStationSouth
			     TrackingStationHub VAB VABMainBuilding
			     VABPodMemorial VABRoundTank VABSouthComplex
			     VABTanks) ],
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
		Bop => [ qw (Poles Slopes Peaks Mara Valley Ridges) ],
		Pol => [ qw (Poles Lowlands Midlands Highlands) ],
		Eeloo => [ qw (Poles Glaciers Midlands Lowlands IceCanyons
			       Highlands Craters) ]
	       );

# Various situations you may find yourself in
my @stockSits = qw (Landed Splashed FlyingLow FlyingHigh InSpaceLow InSpaceHigh);
my @recoSits = qw (Flew FlewBy SubOrbited Orbited Surfaced);
my @scanSits = qw (AltimetryLoRes AltimetryHiRes BiomeAnomaly);

# Reverse-engineered caps for recovery missions.  The values for SubOrbited
# and Orbited are inverted on Kerbin, handled later.
my %recoCap = (
	       Flew => 6,
	       FlewBy => 7.2,
	       SubOrbited => 12,
	       Orbited => 9.6,
	       Surfaced => 18
	      );
# All SCANsat caps are 20
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
my $ksc = 'KSC';

# Only color science if below this threshold
my $threshold = 95;

### Begin!
# Construct sbv hash
while (<DATA>) {
  chomp;
  my @sbvs = split;
  $sbvData{$sbvs[0].$sbvs[1]} = $sbvs[2];
}

# Read in science defs to build prebuild datamatrix for each experiment
open my $defs, '<', "$scidef" or die $ERRNO;
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

    my ($key,$value) = split /=/;
    $key =~ s/\s+//g;		# Clean spaces
    $value =~ s/\s+//g;		# Also fix default spacing in ScienceDefs.cfg

    if ($key eq 'id') {
      @testdef = (@testdef,$value);
    } elsif ($key eq 'situationMask') {
      $value = binary($value);
      @sitmask = (@sitmask,$value);
    } elsif ($key eq 'biomeMask') {
      $value = binary($value);
      @biomask = (@biomask,$value);
    } elsif ($key eq 'requireAtmosphere') {
      @atmo = (@atmo,'1') if $value eq 'True'; # Waiting for fix to sciencedefs
      @atmo = (@atmo,'0') if $value eq 'False';
    } elsif ($key eq 'dataScale') {
      @dataScale = (@dataScale,$value);
    } elsif ($key eq 'scienceCap') {
      @scienceCap = (@scienceCap,$value);
    }
  }
}
close $defs or warn $ERRNO;


## Iterate and decide on conditions, build matrix, gogogo
# Build stock science hash
foreach my $i (0..scalar @testdef - 1) {
  # Array of binary values, only need to do once per test
  my @sits = split //,$sitmask[$i];
  my @bins = split //,$biomask[$i];

  foreach my $planet (0..$planetCount) {
    # Avoid replacing official planet list, thus duplicating Kerbin
    my $stavro = $planets[$planet];
    # Build list of potential situations
    my @situations = @stockSits;

    # Array of arrays for spob-specific biomes, nullify alongside @situations
    my @biomes = ([@{$universe{$stavro}}])x6;
    # KSC biomes are SrfLanded only
    if ($stavro eq $ksc) {
      @situations = qw (Landed);
    } else {
      for (my $binDex = scalar @sits - 1;$binDex>=0;$binDex--) {
	my $zIndex = abs $binDex-5;
	if ($sits[$zIndex] == 0) {
	  splice @situations, $binDex, 1;
	  splice @biomes, $binDex, 1;
	} elsif ($bins[$zIndex] == 0) {
	  $biomes[$binDex] = [ qw (Global)];
	}
      }
    }

    foreach my $sit (0..scalar @situations - 1) {
      # No surface
      next if (($situations[$sit] eq 'Landed') && ($stavro =~ m/^Kerbol$|^Jool$/));
      # Water
      next if (($situations[$sit] eq 'Splashed') && ($stavro !~ m/^Kerbin$|^Eve$|^Laythe$/));
      # Atmosphere
      if ($stavro !~ m/^$ksc$|^Kerbin$|^Eve$|^Duna$|^Jool$|^Laythe$/) {
	next if $situations[$sit] =~ m/^FlyingLow$|^FlyingHigh$/;
	next if $atmo[$i] == 1;
      }
      # Fold KSC into Kerbin, if need be
      # Inconvenient, ruined by the cleaning function later
      if ($stavro eq $ksc && $opt{ksckerbin}) {
	$stavro = 'Kerbin';
      }

      foreach my $bin (0..scalar @{$biomes[$sit]} - 1) {
	# Use specific data (test, spob, sit, biome) as key to allow specific
	# references and unique overwriting
	my $sbVal = $sbvData{$stavro.$situations[$sit]};
	my $cleft = $sbVal*$scienceCap[$i];
	$dataMatrix{$testdef[$i].$stavro.$situations[$sit].$biomes[$sit][$bin]} = [$testdef[$i],$stavro,$situations[$sit],$biomes[$sit][$bin],$dataScale[$i],'1',$sbVal,'0',$cleft,$cleft,'0'];
      }
    }
  }
}

# This is awkwardly saddled between the above and below, but it keeps everyone
# running smoothly in the case of -k
if ($opt{ksckerbin}) {
  splice @planets, 1, 1;
  $planetCount--;
}

# Build recovery hash
foreach my $planet (0..$planetCount) {
  next if $planets[$planet] eq $ksc;

  my @situations = @recoSits;
  shift @situations;		# Either Flew or FlewBy, not both
  # Kerbin is special of course
  if ($planets[$planet] eq 'Kerbin') {
    $situations[0] = 'Flew';	# No FlewBy
    pop @situations;		# No Surfaced
  }

  foreach my $sit (0..scalar @situations - 1) {
    # No surface
    next if (($situations[$sit] eq 'Surfaced') && ($planets[$planet] =~ m/^Kerbol|^Jool/));

    my $sbVal = $sbvData{$planets[$planet].'Recovery'};
    my $cleft;

    # Kerbin's values for (sub)orbital recovery are inverted elsewhere, since
    # you're coming the other way.  Probably a neater way to do this.
    if ($planets[$planet] eq 'Kerbin' && $situations[$sit] eq 'Orbited') {
      $cleft = $sbVal*$recoCap{'SubOrbited'};
    } elsif ($planets[$planet] eq 'Kerbin' && $situations[$sit] eq 'SubOrbited') {
      $cleft = $sbVal*$recoCap{'Orbited'};
    } else {
      $cleft = $sbVal*$recoCap{$situations[$sit]};
    }

    $reco{$planets[$planet].$situations[$sit]} = [$recovery,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
  }
}

# Build SCANsat hash
if ($opt{scansat}) {
  foreach my $planet (0..$planetCount) {
    # No scanning for KSC biomes
    # But *technically* you can scan Jool and Kerbol
    next if ($planets[$planet] eq $ksc);

    my @situations = @scanSits;
    foreach my $sit (0..scalar @situations - 1) {

      my $sbVal = $sbvData{$planets[$planet].'InSpaceHigh'};
      my $cleft = $sbVal*$scanCap;

      # SCANsat results from Jool and Kerbol are reduced by half
      # https://github.com/S-C-A-N/SCANsat/issues/125
      $cleft /= 2 if $planets[$planet] =~ m/^Kerbol$|^Jool$/;

      $scan{$planets[$planet].$situations[$sit]} = [$scansat,$planets[$planet],$situations[$sit],'1','1',$sbVal,'0',$cleft,$cleft,'0'];
    }
  }
}

open my $file, '<', "$pers" or die $ERRNO;
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
    my ($key,$value) = split /=/;
    $key =~ s/\s+//g;		# Clean spaces
    $value =~ s/\s+//g;
    $value =~ s/Sun/Kerbol/g;

    if ($key eq 'id') {
      # Replace recovery and SCANsat data here, why not?
      if ($value =~ m/^$recovery/) {
	$recoTicker = 1;
	$value =~ s/(Flew[By]?|SubOrbited|Orbited|Surfaced)/\@$1/g;
	@pieces = (split /@/, $value);
      } elsif ($opt{scansat} && $value =~ m/^$scansat/) {
	$scanTicker = 1;
	$value =~ s/^$scansat(.*)\@(.*)InSpaceHighsurface$/$scansat\@$2\@$1/g;
	@pieces = (split /@/, $value);
      } else {
	($recoTicker,$scanTicker) = (0,0);
	# Watch out for srf landed/splashed, InSpaceHigh/Low, FlyingHigh/Low
	$value =~ s/Srf(Landed|Splashed)/\@$1\@/g;
	$value =~ s/(InSpace|Flying)(Low|High)/\@$1$2\@/g;
	@pieces = (split /@/, $value);
      }

      # Ensure arrays are the same length
      push @test, $pieces[0];
      push @spob, $pieces[1];
      push @where, $pieces[2];
      push @biome, $pieces[3] // 'Global'; # global biomes
    } elsif ($key =~ m/^title/) {
      @title = (@title,$value);
    } elsif ($key =~ m/^dsc/) {
      @dsc = (@dsc,$value);
    } elsif ($key =~ m/^scv/) {
      @scv = (@scv,$value);
    } elsif ($key =~ m/^sbv/) {
      @sbv = (@sbv,$value);
    } elsif ($key =~ m/^sci/) {
      @sci = (@sci,$value);
    } elsif ($key =~ m/^cap/) {
      @cap = (@cap,$value);

      # Build recovery and SCANsat data hashes
      if ($recoTicker == 1) {
	my $percL = calcPerc($sci[-1],$cap[-1]);
	$reco{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$percL];
      } elsif ($opt{scansat} && $scanTicker == 1) {
	my $percL = calcPerc($sci[-1],$cap[-1]);
	$scan{$pieces[1].$pieces[2]} = [$pieces[0],$pieces[1],$pieces[2],$dsc[-1],$scv[-1],$sbv[-1],$sci[-1],$cap[-1],$cap[-1]-$sci[-1],$percL];
      }
    }
  }
}
close $file or warn $ERRNO;

# Build the matrix
foreach (0..scalar @test - 1) {
  # Exclude tests stored in separate hashes
  next if $test[$_] =~ m/^$scansat|^$recovery/;
  if ($biome[$_]) {
    my $percL = calcPerc($sci[$_],$cap[$_]);

    if ($biome[$_] =~ m/^$ksc|^Runway|^LaunchPad|^VAB|^SPH|^R&D|^Astronaut|^FlagPole|^Mission|^Tracking|^Crawler|^Administration/) {
      # KSC biomes *should* be SrfLanded-only, this ensures that we skip any
      # anomalous data in persistent.sfs.  This complements the test below
      # but saves some work given the KSC/Kerbin potential with the -k flag
      next if $where[$_] ne 'Landed';
      # Take KSC out of Kerbin
      if (!$opt{ksckerbin}) {
	$spob[$_] = $ksc;
      }
    }

    # Skip over annoying "fake" science expts caused by ScienceAlert
    # For more info see
    # http://forum.kerbalspaceprogram.com/threads/76793-0-90-ScienceAlert-1-8-4-Experiment-availability-feedback-%28December-23%29?p=1671187&viewfull=1#post1671187
    next if !$dataMatrix{$test[$_].$spob[$_].$where[$_].$biome[$_]};

    $dataMatrix{$test[$_].$spob[$_].$where[$_].$biome[$_]} = [$test[$_],$spob[$_],$where[$_],$biome[$_],$dsc[$_],$scv[$_],$sbv[$_],$sci[$_],$cap[$_],$cap[$_]-$sci[$_],$percL];
  }
}

###
### Begin the printing process!
###
my @header = qw [Experiment Spob Condition dsc scv sbv sci cap Sci.Left Perc.Accom];
dataSplice(\@header) if !$opt{moredata};

## Prepare fancy-schmancy Excel workbook
# Globals defined here so -e flag works properly
my ($workbook,$bold,$bgRed,$bgGreen);

if (!$opt{excludeexcel}) {
  # Create new workbook
  $workbook = Excel::Writer::XLSX->new( "$outfile" );
  # Bold for headers, red for science left, green for stupidly small values
  $bold = $workbook->add_format();
  $bgRed = $workbook->add_format();
  $bgGreen = $workbook->add_format();

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
  columnWidths($workVars{$recov}[0],9.17,6.5,9);

  if ($opt{scansat}) {
    $workVars{$scansat} = [$workbook->add_worksheet( 'SCANsat' ), 1];
    $workVars{$scansat}[0]->write( 0, 0, \@header, $bold );

    # SCANsat widths, manually determined
    columnWidths($workVars{$scansat}[0],9.17,6.5,11.83);
  }
}

$header[1] = 'Condition';
$header[2] = 'Biome';

if (!$opt{excludeexcel}) {
  foreach my $planet (0..$planetCount) {
    # Interpolate via " instead of '
    $workVars{$planets[$planet]} = [$workbook->add_worksheet( "$planets[$planet]" ), 1];
    $workVars{$planets[$planet]}[0]->write( 0, 0, \@header, $bold );

    # Stock science widths, manually determined
    columnWidths($workVars{$planets[$planet]}[0],15.5,9.67,8.5);
  }
}


## Actually print everybody!
open my $csvOut, '>', "$csvFile" or die $ERRNO if $opt{csv};
writeToCSV(\@header) if $opt{csv};

# Stock science
foreach my $key (sort sitSort keys %dataMatrix) {
  # Splice out planet name so it's not repetitive
  my $planet = splice @{$dataMatrix{$key}}, 1, 1;
  dataSplice(\@{$dataMatrix{$key}}) if !$opt{moredata};
  writeToExcel($planet,\@{$dataMatrix{$key}},$key,\%dataMatrix) if !$opt{excludeexcel};

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
  dataSplice(\@{$reco{$key}}) if !$opt{moredata};
  writeToExcel($recov,\@{$reco{$key}},$key,\%reco) if !$opt{excludeexcel};
  writeToCSV(\@{$reco{$key}}) if $opt{csv};

  if ($opt{tests}) {
    # Neater spacing in test averages output
    buildScienceData($key,$recovery,\%testData,\%reco);
  } elsif ($opt{average}) {
    buildScienceData($key,$recov,\%spobData,\%reco);
  }
}
# SCANsat
if ($opt{scansat}) {
  foreach my $key (sort { specialSort($a, $b, \%scan) } keys %scan) {
    dataSplice(\@{$scan{$key}}) if !$opt{moredata};
    writeToExcel($scansat,\@{$scan{$key}},$key,\%scan) if !$opt{excludeexcel};
    writeToCSV(\@{$scan{$key}}) if $opt{csv};

    if ($opt{tests}) {
      # Neater spacing in test averages output
      buildScienceData($key,$scansatMap,\%testData,\%scan);
    } elsif ($opt{average}) {
      buildScienceData($key,$scansat,\%spobData,\%scan);
    }
  }
}
close $csvOut or warn $ERRNO if  $opt{csv};


## Sorting of different average tables
open my $avgOut, '>', "$avgFile" or die $ERRNO if  $opt{outputavgtable};
# Ensure the -t flag supersedes -a if both are given
if ($opt{average} || $opt{tests}) {
  my $string = "Average science left:\n\n";
  my ($hashRef,$arrayRef);

  if ($opt{tests}) {
    $string .= "Test\t";
    $hashRef = \%testData;
    $arrayRef = \@testdef if !$opt{scienceleft};
  } elsif ($opt{average}) {
    $string .= 'Spob';
    $hashRef = \%spobData;
    $arrayRef = \@planets if !$opt{scienceleft};
  }

  $string .= "\tAvg/exp\tTotal\tCompleted\n";
  print "$string";
  print $avgOut "$string"  if  $opt{outputavgtable};

  if ($opt{percentdone}) {
    average3($hashRef);
  } elsif ($opt{scienceleft}) {
    average2($hashRef);
  } else {
    average1($hashRef,$arrayRef);
  }
}
close $avgOut or warn $ERRNO if  $opt{outputavgtable};


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

    return;
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

# Determine column header widths
sub columnWidths
  {
    my ($sheet,$col1,$col2,$col3) = @_;

    $sheet->set_column( 0, 0, $col1 );
    $sheet->set_column( 1, 1, $col2 );
    $sheet->set_column( 2, 2, $col3 );

    return;
  }

## Custom sort order, adapted from:
## http://stackoverflow.com/a/8171591/2521092
# Kerbin, KSC, and its moons come first, then Kerbol, then proper sorting of
# conditions
# matches worksheets
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


# Properly splice data
sub dataSplice
  {
    my $rowRef = shift;
    splice @{$rowRef}, 3, 3;
    return;
  }

sub writeToCSV
  {
    my $rowRef = shift;

    print $csvOut join q{,} , @{$rowRef};
    print $csvOut "\n";
    return;
  }

sub writeToExcel
  {
    my ($sheetName,$rowRef,$matrixKey,$hashRef) = @_;

    $workVars{$sheetName}[0]->write_row( $workVars{$sheetName}[1], 0, $rowRef );
    if ($opt{moredata}) {
      $workVars{$sheetName}[0]->write( $workVars{$sheetName}[1], 8, ${$hashRef}{$matrixKey}[8], $bgRed ) if ${$hashRef}{$matrixKey}[9] < $threshold;
      $workVars{$sheetName}[0]->write( $workVars{$sheetName}[1], 4, ${$hashRef}{$matrixKey}[4], $bgGreen ) if ((${$hashRef}{$matrixKey}[4] < 0.001) && (${$hashRef}{$matrixKey}[4] > 0));
    } else {
      $workVars{$sheetName}[0]->write( $workVars{$sheetName}[1], 5, ${$hashRef}{$matrixKey}[5], $bgRed ) if ${$hashRef}{$matrixKey}[6] < $threshold;
    }

    $workVars{$sheetName}[1]++;
    return;
  }

# Build data hashes for averages
sub buildScienceData
  {
    my ($key,$ind,$dataRef,$hashRef) = @_;

    # Sci, count, cap
    if ($opt{moredata}) {
      ${$dataRef}{$ind}[0] += ${$hashRef}{$key}[8];
      ${$dataRef}{$ind}[2] += ${$hashRef}{$key}[7];
    } else {
      ${$dataRef}{$ind}[0] += ${$hashRef}{$key}[5];
      ${$dataRef}{$ind}[2] += ${$hashRef}{$key}[4];
    }
    ${$dataRef}{$ind}[1]++;

    return;
  }


# Alphabeticalish averages
sub average1
  {
    my $hashRef = shift;
    my $arrayRef = shift;

    if ($opt{tests}) {
      push @{$arrayRef}, $recovery; # Neater spacing in test averages output
      push @{$arrayRef}, $scansatMap if $opt{scansat};
      @{$arrayRef} = sort @{$arrayRef};
    }

    foreach my $index (0..scalar @{$arrayRef} - 1) {
      printAverageTable(${$arrayRef}[$index],$hashRef);
    }

    if (!$opt{tests}) {
      printAverageTable($recov,$hashRef);
      printAverageTable($scansat,$hashRef) if $opt{scansat};
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

    my $indShort = substr $ind, 0, 14; # Neater spacing in test averages output
    my $avg = $hash{$ind}[0]/($hash{$ind}[1]);
    my $remains = $hash{$ind}[2] - $hash{$ind}[0];
    my $per = 100*$remains/$hash{$ind}[2];

    printf "%s\t%.0f\t%.0f\t%.0f\n", $indShort, $avg, $hash{$ind}[0], $per;
    if ($opt{outputavgtable}) {
      printf $avgOut "%s\t%.0f\t%.0f\t%.0f\n", $ind, $avg, $hash{$ind}[0], $per;
    }

    return;
  }

#### Usage statement ####
# Escapes not necessary but ensure pretty colors
# Final line must be unindented?
sub usage
  {
    print <<USAGE;
Usage: $PROGRAM_NAME [-atspikmcneo -h -f path/to/dotfile ]
       $PROGRAM_NAME [-g <game_location> -u <savefile_name>]

       $PROGRAM_NAME [-ATSPIKMCNEO -G -U] -> Turn off a given option

      -a Display average science left for each planet
      -t Display average science left for each experiment type.  Supersedes -a.

      -s Sort by science left, including output file(s) and averages from -a
         and -t flags
      -p Sort by percent science accomplished, including output file(s) and
         averages from -a and -t flags.  Supersedes -s.

      -i Include data from SCANsat
      -k List data from KSC biomes as being from Kerbin (same Excel worksheet)
      -m Add some largely boring data to the output (i.e., dsc, sbv, scv)
      -c Output data to csv file as well
      -n Turn off formatted printing in Excel (i.e., colors and bolding)
      -e Don't output the Excel file
      -o Save the chosen average table to a file.  Requires -a or -t.

      -g Specify path to your KSP folder
      -u Enter the username of your KSP save folder; otherwise, whatever files
         are present in the local directory will be used.
      -f Specify path to config file.  Supersedes a local .parsesciencerc file.
      -h Print this message
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
