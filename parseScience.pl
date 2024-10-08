#!/usr/bin/env perl
# parseScience.pl by Amory Meltzer
# v0.97.3
# https://github.com/Amorymeltzer/ksp
# Parse a KSP persistent.sfs file, report science information
# Leftover science in red, candidates for manual cleanup in green
# Sun represented as Kerbol
# MajorCrater triplication hack is dirty but that's on Squad!
# CresentBay actually real?  Doubtful but giving Squad benefit of the doubt

### FIXES, TODOS
### UPDATE FOR 1.1
## asteroid and cometSamples: can they be landed at KSC biomes??
## Can you do srfsplashed in every biome on other planets with water?
## Incorporate multiplier?  Might look weird...
## Biome sort incorporated better?  Elsewhere?  With -a or -t options?
## Report data test by condition?
## Let report be not stock-only
## -a/-t Always print, what if just want report?  Non-parallel behavior with
## -a/-t printing and -o/-r file output?
### Add -r report handling to buildScienceData??
## Default windows/linux path to Gamedata/pers/scidefs/etc.?
### Steam locations
## Flying high/low at Sun, see also multipliers
### Game might think possible but it's not actually real...

## Cleanup data/test hashes, the order of the data is unintuitive

use 5.010;

use English;
use autodie qw(open close);

use strict;
use warnings;
# use diagnostics;

use Getopt::Std;

# Parse command line options
my %opts = ();
getopts('aAtTbBsSpPiIjJdDkKlLmMzZcCnNeEoOrRg:Gu:Uf:h', \%opts);

if ($opts{h}) {
  usage();
  exit;
}

# Correlate commandline flags with their corresponding config options
my %lookup = (g => 'gamelocation',
	      u => 'username',
	      a => 'average',
	      t => 'tests',
	      b => 'biome',
	      s => 'scienceleft',
	      p => 'percentdone',
	      i => 'scansat',
	      j => 'ignoreasteroids',
	      d => 'breakingground',
	      k => 'ksckerbin',
	      l => 'noksc',
	      m => 'moredata',
	      z => 'unfinishedonly',
	      c => 'csv',
	      n => 'noformat',
	      e => 'excludeexcel',
	      o => 'outputavgtable',
	      r => 'report'
	     );

### ENVIRONMENT VARIABLES
## Build %opt hash, using above lookup table.  Replaced from the %opts table.
my %opt;
foreach my $key (keys %lookup) {
  $opt{$lookup{$key}} = 0;
}


# Figure out where this script is, for proper reading/writing of files
my $scriptDir;                  # Directory of this script
# use Cwd 'abs_path';
use Cwd qw(abs_path cwd);
use File::Basename 'dirname';
# Other useful shorthands for finding files
my $rc   = 'parsesciencerc';    # Config dotfile name of choice.  Wordy.
my $cwd  = cwd();               # Current working directory
my $home = $ENV{HOME};          # MAGIC hash with user env variables for $home

BEGIN {
  $scriptDir = dirname abs_path __FILE__;
}

## .parsesciencerc config file
my $dotfile;

# Round up the usual suspects, all superseded by commandline flag
my @dotLocales = ("$cwd/.$rc", "$scriptDir/.$rc");
# Windows (XP anyway) complains about $home
@dotLocales = (@dotLocales, "$home/.$rc", "$home/.config/parseScience/$rc") if $home;
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
  open my $dot, '<', "$dotfile";
  while (<$dot>) {
    chomp;

    next if m/^#/g;    # Ignore comments
    next if !$_;       # Ignore blank lines

    if (!m/^\w+ = \w+/) {    # Ignore and warn on malformed entries
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
  close $dot;
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


### FILE DEFINITIONS
my ($scidef, $pers, $breakingScidef, $path);
my $scidefName = 'ScienceDefs.cfg';
my $persName   = 'persistent.sfs';
my $gdsr       = 'GameData/Squad/Resources/';
my $serenityr  = 'GameData/SquadExpansion/Serenity/Resources/';

# Build and iterate through all potential options
my @scidefLocales = ("$cwd/$scidefName", "$scriptDir/$scidefName");
my @persLocales   = ("$cwd/$persName",   "$scriptDir/$persName");

if (!$opt{gamelocation}) {
  if ($OSNAME eq 'darwin') {
    $path = '/Applications/KSP_osx/';
  } elsif ($OSNAME eq 'linux') {
    $path = '/Applications/KSP_linux/';
  } elsif ($OSNAME eq 'MSWin32') {
    $path = 'C:/Program Files/KSP-win/';
  }
} else {
  $path          = $opt{gamelocation};
  @scidefLocales = ($path.$gdsr.$scidefName, @scidefLocales);
}

if ($opt{username}) {
  @scidefLocales = ($path.$gdsr.$scidefName, @scidefLocales);
  @persLocales   = ($path."saves/$opt{username}/".$persName, @persLocales);
}

# Test files for existance
$scidef = checkFiles($scidef, $scidefName, \@scidefLocales);
$pers   = checkFiles($pers,   $persName,   \@persLocales);

if ($opt{breakingground}) {
  # Can't process multiple files with the same name in the sample place
  warnNicely('breakingground option selected but no username or gamelocation given (-u or -g)', 1) if (!$opt{username} && !$opt{gamelocation});

  $breakingScidef = checkFiles($breakingScidef, $scidefName, [$path.$serenityr.$scidefName]);
}



# Don't bother outputting average file if there ain't any averages to save
if (!$opt{average} && !$opt{tests}) {
  if ($opt{outputavgtable}) {
    $opt{outputavgtable} = 0;
    warnNicely('outputavgtable option given but no data table selected (-a or -t).  Skipping...');
  }
  if ($opt{report}) {
    $opt{report} = 0;
    warnNicely('report option given but no data table selected (-a or -t).  Skipping...');
  }
}

# Only load if necessary
if (!$opt{excludeexcel}) {
  require Excel::Writer::XLSX;
}

my $outfile = 'scienceToDo.xlsx';
my $csvFile = 'scienceToDo.csv';
my $avgFile = 'average_table.txt';
my $rptFile = 'report.csv';

### GLOBAL VARIABLES
my %dataMatrix;    # Stock data
my %reco;          # Craft recovery data
my %scan;          # SCANsat data
my %sbvData;       # sbv values from END data
my %workVars;      # Hash of arrays to hold worksheets, current row
my %spobData;      # Science per spob
my %testData;      # Science per test
my %report;        # Hold basic report data

# ScienceDefs.cfg variables
my (@testdef,      # Basic test names
    @sitmask,      # Where test is valid
    @biomask,      # Where biomes for test matter
    @atmo,         # Check if atmosphere required or not
    @noAtmo,       # Check if NO atmosphere required or not
    @dataScale,    # dataScale, same as dsc in persistent.sfs
    @scienceCap    # Base experiment cap, multiplied by sbv
   );

# persistent.sfs variables
my (@dsc,          # Data scale
    @scv,          # Percent left to research
    @sbv,          # Base value multiplier to reach cap
    @sci,          # Science researched so far
    @cap           # Max science
   );

# Store details from split id
my @pieces;
my (@test,         # Which test
    @spob,         # Which planet/moon
    @where,        # What activity
    @biome         # What biome
   );

my @planets = qw (Kerbin KSC Mun Minmus Kerbol Moho Eve Gilly Duna Ike Dres
  Jool Laythe Vall Tylo Bop Pol Eeloo);

# Different spobs, different biomes
my %universe = (Kerbin => [qw (Badlands Deserts Grasslands Highlands IceCaps
			   Mountains NorthernIceShelf Shores SouthernIceShelf
			   Tundra Water)],
		KSC => [qw (KSC Administration AstronautComplex Crawlerway
			FlagPole LaunchPad MissionControl Runway SPH
			SPHMainBuilding SPHRoundTank SPHTanks SPHWaterTower
			R&DCentralBuilding R&DCornerLab R&DMainBuilding
			R&DObservatory R&DSideLab R&DSmallLab R&DTanks
			R&DWindTunnel TrackingStation
			TrackingStationDishEast TrackingStationDishNorth
			TrackingStationDishSouth TrackingStationHub VAB
			VABMainBuilding VABPodMemorial VABRoundTank
			VABSouthComplex VABTanks IslandAirfield Baikerbanur
			BaikerbanurLaunchpad)],
		# Northeast Basin is displayed, but listed as Northern Basin
		Mun => [qw (Canyons EastCrater EastFarsideCrater FarsideBasin
			FarsideCrater HighlandCraters Highlands Lowlands
			MidlandCraters Midlands NorthernBasin
			NorthwestCrater PolarCrater PolarLowlands Poles
			SouthwestCrater TwinCraters)],
		Minmus => [qw (Flats GreatFlats GreaterFlats Highlands
			   LesserFlats Lowlands Midlands Poles Slopes)],
		Kerbol => [qw (Global)],
		Moho   => [qw (Canyon CentralLowlands Highlands Midlands
			 MinorCraters NorthPole NorthernSinkhole
			 NorthernSinkholeRidge SouthEasternLowlands
			 SouthPole SouthWesternLowlands WesternLowlands)],
		Eve => [qw (AkatsukiLake CraterLake Craters EasternSea
			ExplodiumSea Foothills Highlands ImpactEjecta
			Lowlands Midlands Olympus Peaks Poles Shallows
			WesternSea)],
		Gilly => [qw (Highlands Lowlands Midlands)],
		Duna  => [qw (Craters EasternCanyon Highlands Lowlands
			 MidlandCanyon MidlandSea Midlands NortheastBasin
			 NorthernShelf PolarCraters PolarHighlands Poles
			 SouthernBasin WesternCanyon)],
		Ike => [qw (CentralMountainRidge EasternMountainRidge Lowlands
			Midlands PolarLowlands SouthEasternMountainRange
			SouthPole WesternMountainRidge)],
		Dres => [qw (Canyons Highlands ImpactCraters ImpactEjecta
			 Lowlands Midlands Poles Ridges)],
		Jool   => [qw (Global)],
		Laythe => [qw (CraterBay CraterIsland CresentBay DegrasseSea
			   Dunes Peaks Poles Shallows Shores TheSagenSea)],
		Vall => [qw (Highlands Lowlands Midlands Mountains
			 NortheastBasin NorthwestBasin Poles SouthernBasin
			 SouthernValleys)],
		Tylo => [qw (GagarinCrater GalileioCrater GrissomCrater
			 Highlands Lowlands Mara Midlands MinorCraters
			 TychoCrater)],
		Bop   => [qw (Peaks Poles Ridges Slopes Valley)],
		Pol   => [qw (Highlands Lowlands Midlands Poles)],
		Eeloo => [qw (BabbagePatch Craters Fragipan Highlands IceCanyons
			  Lowlands Midlands Mu NorthernGlaciers Poles
			  SouthernGlaciers)]
	       );

# Various situations you may find yourself in
my @stockSits = qw (Landed Splashed FlyingLow FlyingHigh InSpaceLow InSpaceHigh);
my @recoSits  = qw (Flew FlewBy SubOrbited Orbited Surfaced);
my @scanSits  = qw (AltimetryLoRes AltimetryHiRes BiomeAnomaly Resources Visual);

# Stupid resource scan earns science for each spob, but *isn't* a science
# definition, so we need to look it up by PlanetId.  Toss an undef in there to
# make it easier.
my @planetID = qw(Kerbin Mun Minmus Moho Eve Duna Ike Jool Laythe Vall Bop Tylo Gilly Pol Dres Eeloo);
unshift @planetID, undef;


# Some lookup hashes
my %noLandLookup     = makeMap([qw (Kerbol Jool)]);
my %waterLookup      = makeMap([qw (Kerbin Eve Laythe)]);
my %kscLookup        = makeMap($universe{KSC});
my %atmosphereLookup = makeMap([$universe{KSC}->@*, qw (Kerbin Eve Duna Jool Laythe)]);

# Help speed up sorting
my %memoized_situation = ();
my %memoized_spob      = ();

# Common regexes used in specialSort
my $SPOB_RE        = join q{|}, @planets;
my %spec_order_map = map {$planets[$_] => $_} 0 .. $#planets;
my @condOrder      = (@recoSits, @scanSits);
my %cond_order_map = map {$condOrder[$_] => $_} 0 .. $#condOrder;
my $COND_RE        = join q{|}, @condOrder;
# Common regex used in sitSort
my $SIT_RE        = join q{|}, @stockSits;
my %sit_order_map = map {$stockSits[$_] => $_} 0 .. $#stockSits;
# Common regex for finding science loops.  The Breaking Ground expansion
# SciencesDefs.cfg has some weird characters, so we have to use a regex rather
# than matching the line exactly, which slows us down quite a bit.
my $SCI_RE = '^EXPERIMENT_DEFINITION';
# Regex for finding Breaking Ground spob-specific rover science
my $ROC_RE = '^ROCScience_';

# Reverse-engineered caps for recovery missions.  The values for SubOrbited
# and Orbited are inverted on Kerbin, handled later.
my %recoCap = (Flew       => 6,
	       FlewBy     => 7.2,
	       SubOrbited => 12,
	       Orbited    => 9.6,
	       Surfaced   => 18
	      );
# All SCANsat caps are 20 FIXME TODO CHECK THIS
my $scanCap = 20;

# Am I in a science, recovery, or SCANsat loop?
my ($ticker, $recoTicker, $scanTicker) = (0, 0, 0);

# Sometimes I use one versus the other, mainly for spacing in averages table
my $recov      = 'Recov';
my $recovery   = 'recovery';
my $scansat    = 'SCANsat';
my $scansatMap = 'SCANsatMapping';
my $ksc        = 'KSC';
my $total      = 'total';

# Only color science if below this percentage
my $colorThreshold = 95;
# Index for printing out data
my $dataIdx = $opt{moredata} ? 8 : 5;
# Excel column widths, manually determined
## no critic (ProhibitMagicNumbers)
my %columnSizes = ($recov   => [9.17, 6.5,  9],
		   $scansat => [9.17, 6.5,  11.83],
		   spob     => [15.5, 9.67, 8.5]
		  );
## use critic
# Construct sbv hash
%sbvData = map {my ($spob, $sit, $val) = split; $spob.$sit => $val} <DATA>;

### Begin!
# Read in science defs to build prebuild datamatrix for each experiment
open my $defs, '<', "$scidef";
readSciDefs($defs);
close $defs;
if ($opt{breakingground}) {
  open my $bdefs, '<', "$breakingScidef";
  readSciDefs($bdefs);
  close $bdefs;
}


my ($rocSpob, $rocName);
my @tempPlanets;
## Iterate and decide on conditions, build matrix, gogogo
# Build stock science hash
foreach my $i (0 .. $#testdef) {
  next if ($testdef[$i] =~ /^asteroid|^infrared|^cometS/ && $opt{ignoreasteroids});

  # Deal with planet-specific science from Breaking Ground, and avoid replacing
  # the official planet list
  if ($testdef[$i] =~ /$ROC_RE/) {
    ($rocSpob, $rocName) = $testdef[$i] =~ /$ROC_RE($SPOB_RE)(.+)/;
    @tempPlanets = ($rocSpob);
    # Give it a good name
    $testdef[$i] = $rocName;
  } else {
    @tempPlanets = @planets;
    if ($opt{noksc}) {
      splice @tempPlanets, 1, 1;
    }
  }

  # Array of binary values, only need to do once per test
  my @sits = split //, $sitmask[$i];
  my @bins = split //, $biomask[$i];

  foreach my $planet (@tempPlanets) {
    # Build list of potential situations
    my @situations = @stockSits;
    # Array of arrays for spob-specific biomes, nullify alongside @situations
    my @biomes = ([@{$universe{$planet}}]) x 6;

    # KSC biomes are SrfLanded only
    if ($planet eq $ksc) {
      next if $bins[-1] == 0;
      @situations = qw (Landed);
    } else {
      # Would be good to include something here that handles 0 instead of 1..63,
      # mainly because SCANsat does that.  Annoyingly.  FIXME TODO
      for my $binDex (reverse 0 .. $#sits) {
	my $zIndex = abs $binDex - 5;
	if ($sits[$zIndex] == 0) {
	  splice @situations, $binDex, 1;
	  splice @biomes,     $binDex, 1;
	} elsif ($bins[$zIndex] == 0) {
	  $biomes[$binDex] = [qw (Global)];
	}
      }
    }

    foreach my $sit (0 .. $#situations) {
      # Water
      next if (($situations[$sit] eq 'Splashed') && (!$waterLookup{$planet}));
      # Atmosphere
      if (!$atmosphereLookup{$planet}) {
	next if ($situations[$sit] eq 'FlyingLow' || $situations[$sit] eq 'FlyingHigh');
	next if $atmo[$i] eq 'True';
      } elsif ($noAtmo[$i] eq 'True') {
	next;
      }
      # No surface
      next if (($situations[$sit] eq 'Landed') && ($noLandLookup{$planet}));
      # Fold KSC into Kerbin, if need be
      # Inconvenient, ruined by the cleaning function later
      if ($planet eq $ksc && $opt{ksckerbin}) {
	$planet = 'Kerbin';
      }

      foreach my $bin (0 .. $#{$biomes[$sit]}) {
	# Use specific data (test, spob, sit, biome) as key to allow specific
	# references and unique overwriting
	my $sbVal = $sbvData{$planet.$situations[$sit]};
	my $cleft = $sbVal * $scienceCap[$i];
	$dataMatrix{$testdef[$i].$planet.$situations[$sit].$biomes[$sit][$bin]} = [$testdef[$i], $planet, $situations[$sit], $biomes[$sit][$bin], $dataScale[$i], '1', $sbVal, '0', $cleft, $cleft, '0'];
      }
    }

    # Don't go applying planet-specific science from Breaking Ground, especially
    # now that we've renamed the test names
    last if $rocSpob;
  }
}

# Manually add resource scans.  Can this be part of the above loop? FIXME TODO
foreach my $planet (@planets) {
  # No scanning KSC or the sun
  next if ($planet eq $ksc || $planet eq 'Kerbol');

  # Only one!  Annoyingly, resource scans make sense as InSpaceHigh, but Kerbin
  # is a holdout: every spob has Recovery matching InSpaceHigh, except Kerbin,
  # where it matches InSpaceLow.  So, technically, the situation *value* is
  # Recovery, but I'll be using InSpaceHigh for meaningfulness.
  my $sit = 'Recovery';

  # ScienceCap is 10
  my $sbVal = $sbvData{$planet.'Recovery'};
  my $cleft = $sbVal * 10;

  # Copied datascale from SCANsat science, but maybe that's not right?  It's the
  # size of what is transmitted, which isn't shown... FIXME TODO
  $dataMatrix{'resourceScan'.$planet.'InSpaceHigh'.'Global'} = ['resourceScan', $planet, 'InSpaceHigh', 'Global', 2, 0, $sbVal, '0', $cleft, $cleft, '0'];
}

# This is awkwardly saddled between the above and below, but it keeps everyone
# running smoothly in the case of -k or -l
# CAN THIS BE MOVED TO JUST ABOVE PERS FIXME TODO
if ($opt{ksckerbin} || $opt{noksc}) {
  splice @planets, 1, 1;
}

# Combine these two?? FIXME TODO
# Build recovery hash
foreach my $planet (@planets) {
  next if $planet eq $ksc;

  # Either Flew or FlewBy, not both
  my @situations = @recoSits[1 .. $#recoSits];

  # Kerbin is special of course
  if ($planet eq 'Kerbin') {
    $situations[0] = 'Flew';    # No FlewBy
    pop @situations;            # and no surfaced
  }
  # No Surfaced
  if ($noLandLookup{$planet}) {
    pop @situations;
  }


  foreach my $sit (0 .. $#situations) {
    my $sbVal = $sbvData{$planet.'Recovery'};
    my $cleft;

    # Kerbin's values for (sub)orbital recovery are inverted elsewhere, since
    # you're coming the other way.  Probably a neater way to do this.
    if ($planet eq 'Kerbin' && $situations[$sit] eq 'Orbited') {
      $cleft = $sbVal * $recoCap{'SubOrbited'};
    } elsif ($planet eq 'Kerbin' && $situations[$sit] eq 'SubOrbited') {
      $cleft = $sbVal * $recoCap{'Orbited'};
    } else {
      $cleft = $sbVal * $recoCap{$situations[$sit]};
    }

    $reco{$planet.$situations[$sit]} = [$recovery, $planet, $situations[$sit], '1', '1', $sbVal, '0', $cleft, $cleft, '0'];
  }
}

# Build SCANsat hash
if ($opt{scansat}) {
  foreach my $planet (@planets) {
    # No scanning for KSC biomes
    # But *technically* you can scan Jool and Kerbol
    next if ($planet eq $ksc);

    my @situations = @scanSits;
    foreach my $sit (0 .. $#situations) {

      # All SCANsat is just one run of a test, and uses InSpaceHigh as far as I
      # am concerned, but in its own ScienceDefs.cfg the situationMask is 0.
      my $sbVal = $sbvData{$planet.'InSpaceHigh'};
      my $cleft = $sbVal * $scanCap;

      # SCANsat results from Jool and Kerbol are reduced by half
      # https://github.com/S-C-A-N/SCANsat/issues/125
      $cleft /= 2 if $noLandLookup{$planet};

      $scan{$planet.$situations[$sit]} = [$scansat, $planet, $situations[$sit], '1', '1', $sbVal, '0', $cleft, $cleft, '0'];
    }
  }
}


# Read in the science we have!  Woo!
open my $file, '<', "$pers";
readPers($file);
close $file;


###
### Begin the printing process!
###
my @header = qw [Experiment Spob Condition dsc scv sbv sci cap Sci.Left Perc.Accom];
dataSplice(\@header) if !$opt{moredata};

## Prepare fancy-schmancy Excel workbook
# Globals defined here so -e flag works properly
my ($workbook, $bold, $bgRed, $bgGreen);
# Can be combined FIXME TODO
if (!$opt{excludeexcel}) {
  # Create new workbook
  $workbook = Excel::Writer::XLSX->new("$outfile");
  # Bold for headers, red for science left, green for stupidly small values
  $bold    = $workbook->add_format();
  $bgRed   = $workbook->add_format();
  $bgGreen = $workbook->add_format();

  # Turn off formatting if so desired
  if (!$opt{noformat}) {
    $bold->set_bold();
    $bgRed->set_bg_color('red');
    $bgGreen->set_bg_color('green');
  }

  # Generate each worksheet with proper header
  # Subroutine these ;;;;;; ##### FIXME TODO
  $workVars{$recov} = [$workbook->add_worksheet('Recovery'), 1];
  $workVars{$recov}[0]->write(0, 0, \@header, $bold);

  columnWidths($workVars{$recov}[0], $columnSizes{$recov});

  if ($opt{scansat}) {
    $workVars{$scansat} = [$workbook->add_worksheet('SCANsat'), 1];
    $workVars{$scansat}[0]->write(0, 0, \@header, $bold);

    columnWidths($workVars{$scansat}[0], $columnSizes{$scansat});
  }
}

$header[1] = 'Condition';
$header[2] = 'Biome';

if (!$opt{excludeexcel}) {
  foreach my $planet (@planets) {
    # Interpolate via " instead of '
    $workVars{$planet} = [$workbook->add_worksheet("$planet"), 1];
    $workVars{$planet}[0]->write(0, 0, \@header, $bold);

    columnWidths($workVars{$planet}[0], $columnSizes{spob});
  }
}


## Actually print everybody!
open my $csvOut, '>', "$csvFile" if $opt{csv};
writeToCSV(\@header) if $opt{csv};

# Stock science
foreach my $key (sort sitSort keys %dataMatrix) {
  # Splice out planet name so it's not repetitive
  my $planet = splice @{$dataMatrix{$key}}, 1, 1;
  dataSplice(\@{$dataMatrix{$key}})                                if !$opt{moredata};
  writeToExcel($planet, \@{$dataMatrix{$key}}, $key, \%dataMatrix) if (!$opt{excludeexcel} && (!$opt{unfinishedonly} || $dataMatrix{$key}[-1] ne '100.00'));

  if ($opt{tests}) {
    buildScienceData($key, $dataMatrix{$key}[0], \%testData, \%dataMatrix);
    if ($opt{report}) {
      buildReportData($key, $planet, $dataMatrix{$key}[0], \%dataMatrix);
    }
  } elsif ($opt{average}) {
    buildScienceData($key, $planet, \%spobData, \%dataMatrix);
    if ($opt{report}) {
      buildReportData($key, $planet, $dataMatrix{$key}[1], \%dataMatrix);
    }
  }

  # Add in spob name to csv, only necessary for stock science
  $dataMatrix{$key}[1] .= "\@$planet";
  writeToCSV(\@{$dataMatrix{$key}}) if ($opt{csv} && (!$opt{unfinishedonly} || $dataMatrix{$key}[-1] ne '100.00'));
}
# Recovery
processData(\%reco, $recov, $recovery);
# SCANsat
if ($opt{scansat}) {
  processData(\%scan, $scansat, $scansatMap);
}
close $csvOut if $opt{csv};


## Report matrix of some interesting totals
open my $rptOut, '>', "$rptFile" if $opt{report};
if ($opt{tests}) {
  printReportTable(@testdef);
} elsif ($opt{average}) {
  printReportTable(@stockSits);
}
close $rptOut if $opt{report};

## Sorting of different average tables
open my $avgOut, '>', "$avgFile" if $opt{outputavgtable};
# Ensure the -t flag supersedes -a if both are given
if ($opt{average} || $opt{tests}) {
  my $string = "Average science left:\n\n";
  my ($hashRef, $arrayRef);

  if ($opt{tests}) {
    $string .= "Test\t";
    $hashRef  = \%testData;
    $arrayRef = \@testdef if !$opt{scienceleft};
  } elsif ($opt{average}) {
    $string .= 'Spob';
    $hashRef  = \%spobData;
    $arrayRef = \@planets if !$opt{scienceleft};
  }

  $string .= "\tAvg/exp\tTotal\tCompleted\n";
  print "$string";
  print $avgOut "$string" if $opt{outputavgtable};

  if ($opt{percentdone}) {
    averagePercent($hashRef);
  } elsif ($opt{scienceleft}) {
    averageRemaining($hashRef);
  } else {
    averageAlphabetical($hashRef, $arrayRef);
  }
}
close $avgOut if $opt{outputavgtable};


### SUBROUTINES
sub makeMap {
  return map {$_ => 1} $_[0]->@*;
}

sub warnNicely {
  my ($err, $ilynPayne) = @_;
  if ($ilynPayne) {
    print 'ERROR: ';
  } else {
    print 'Warning: ';
  }
  print "$err\n";
  exit if $ilynPayne;

  return;
}

sub checkFiles {
  my ($check, $name, $locRef) = @_;
  my $lastOne = pop @{$locRef};

  if (scalar @{$locRef}) {
    foreach my $place (@{$locRef}) {
      $check = $place;
      if (-e $check) {
	last;
      } else {
	warnNicely("No $name file found at $check");
      }
    }
  } else {
    $check = $lastOne;
  }

  if (!-e $check) {
    warnNicely("No $name file found at $check", 1) if !-e $check;
  }

  return $check;
}

# Read in science defs to prebuild the data matrix for each experiment
sub readSciDefs {
  my $fh = shift;

  while (<$fh>) {
    chomp;

    # Only care about science loops
    next if ($ticker == 0 && !/$SCI_RE/);

    # So find them!
    if (/$SCI_RE/) {
      $ticker = 1;
      next;
    }

    # Note when we close out of a loop, nothing valuable after that, so reset
    # and move on
    if (m/^\tRESULTS/) {
      # Confirm proper length of items in case some definitions are missing a
      # key/value pair; right now, it's just requireNoAtmosphere in Breaking
      # Ground
      if ($#noAtmo < $#testdef) {
	@noAtmo = (@noAtmo, 'False');
      }

      $ticker = 0;
      next;
    }

    # Skip the first line, remove leading tabs, and assign arrays
    if ($ticker == 1) {
      next if m/^[\{\s]+$/;    # Take into account blank lines
      s/^\t//i;

      my ($key, $value) = split /=/;
      # Just process the key for now, as we're gonna skip a few, so we can save
      # processing the value for later.  The whitespace could be moved up, but
      # doesn't really save time.  Should probably just skip everything that
      # *isn't* one of the keys we want. FIXME TODO
      $key =~ s/\s+//g;    # Clean spaces and fix default spacing in ScienceDefs.cfg

      # Unnecessary, unused.  baseValue is NOT the same as sbv later, but seems
      # like it should be!  requiredExperimentLevel is just for surface sample
      next if ($key eq 'title' || $key eq 'baseValue' || $key eq 'requiredExperimentLevel');

      for ($value) {
	s/\s+//g;       # Clean spaces and fix default spacing in ScienceDefs.cfg
	s/\/\/.*//g;    # Remove any comments, currently only magnetometer sitmask
      }

      if ($key eq 'id') {
	@testdef = (@testdef, $value);
      } elsif ($key eq 'situationMask') {
	# evaScience is weird.  It's a part that kerbals use, and they can only do
	# so when landed or in space, but the ScienceDefs.cfg entry lists the
	# situationMask as 63; it should be 49
	if ($testdef[-1] eq 'evaScience') {
	  $value = 49;
	}
	@sitmask = (@sitmask, binary($value));
      } elsif ($key eq 'biomeMask') {
	@biomask = (@biomask, binary($value));
      } elsif ($key eq 'requireAtmosphere') {
	@atmo = (@atmo, $value);
      } elsif ($key eq 'requireNoAtmosphere') {
	@noAtmo = (@noAtmo, $value);
      } elsif ($key eq 'dataScale') {
	@dataScale = (@dataScale, $value);
      } elsif ($key eq 'scienceCap') {
	@scienceCap = (@scienceCap, $value);
      }
    }
  }
}


# Convert string to binary, pad to six digits
sub binary {
  my $ones = sprintf '%b', shift;
  while (length($ones) < 6) {
    $ones = '0'.$ones;
  }
  return $ones;
}

# Read in persistent.sfs save file
sub readPers {
  my $fh = shift;

  while (<$fh>) {
    chomp;

    # Comes after all the science, saves a ton of time in large files
    last if /^\t\tname = VesselRecovery$/;

    # Find all the science loops, and also resource scan data
    if (/^\t\tScience$/ || /^\t\t\tPLANET_SCAN_DATA$/) {
      $ticker = 1;
      next;
    }

    # Note when we close out of a loop
    if (/^\t\t\}$/) {
      $ticker = 0;
      next;
    }

    # Skip the first line, remove leading tabs, and assign arrays
    if ($ticker == 1) {
      next if m/^\t\t\{/;
      s/\s+//g;    # Remove whitespace
      my ($key, $value) = split /=/;

      # Unnecessary, unused
      next if $key eq 'title';

      if ($key eq 'PlanetId') {
	my $spob  = $planetID[$value];
	my $sbVal = $sbvData{$spob.'Recovery'};
	# The "science" is all or nothing, which means we don't need all the
	# values, except of course we're just copying the datascale from SCANsat
	# regardless or whether it's the same or not FIXME TODO
	my $dataValue = ['resourceScan', $spob, 'InSpaceHigh', 'Global', 2, 0, $sbVal, $sbVal * 10, $sbVal * 10, 0, '100.00'];
	my $dataKey   = 'resourceScan'.$spob.'InSpaceHigh'.'Global';
	$dataMatrix{$dataKey} = $dataValue;
	next;
      } elsif ($key eq 'id') {
	$value =~ s/Sun/Kerbol/g;
	# Replace recovery and SCANsat data here, why not?
	if ($value =~ /^$recovery/) {
	  $recoTicker = 1;
	  $value =~ s/(Flew[By]?|SubOrbited|Orbited|Surfaced)/\@$1/g;
	} elsif ($value =~ m/^$scansat/ && $opt{scansat}) {
	  $scanTicker = 1;
	  $value =~ s/^$scansat(.*)\@(.*)InSpaceHigh$/$scansat\@$2\@$1/g;
	} else {
	  # Just in case...
	  ($recoTicker, $scanTicker) = (0, 0);
	  # Watch out for srf landed/splashed, InSpaceHigh/Low, FlyingHigh/Low
	  my $asd = $value =~ s/Srf(Landed|Splashed)/\@$1\@/g;
	  if (!$asd) {
	    $value =~ s/((?:InSpace|Flying)(?:Low|High))/\@$1\@/g;
	  }
	}
	@pieces = (split /@/, $value);

	# Ensure arrays are the same length
	push @test,  $pieces[0];
	push @spob,  $pieces[1];
	push @where, $pieces[2];
	push @biome, $pieces[3] // 'Global';    # global biomes
      } elsif ($key eq 'dsc') {
	@dsc = (@dsc, $value);
      } elsif ($key eq 'scv') {
	@scv = (@scv, $value);
      } elsif ($key eq 'sbv') {
	@sbv = (@sbv, $value);
      } elsif ($key eq 'sci') {
	@sci = (@sci, $value);
      } elsif ($key eq 'cap') {
	@cap = (@cap, $value);

	# Rename Breaking Ground planet-specific tests, as done previously when
	# building the dataMatrix from ScienceDefs.cfg
	if ($test[-1] =~ /$ROC_RE/) {
	  $test[-1] =~ s/$ROC_RE(?:$SPOB_RE)(.+)$/$1/;
	}

	# Build data matrix for each piece.  recovery and SCANsat get their own
	# data hashes, and the main, stock science gets some extra pieces below.
	# First, define some common parts.
	my $percL     = calcPerc($sci[-1], $cap[-1]);
	my $dataKey   = $spob[-1].$where[-1];
	my $dataValue = [$test[-1], $spob[-1], $where[-1], $dsc[-1], $scv[-1], $sbv[-1], $sci[-1], $cap[-1], $cap[-1] - $sci[-1], $percL];

	if ($recoTicker == 1) {
	  $reco{$dataKey} = $dataValue;
	  $recoTicker = 0;
	} elsif ($opt{scansat} && $scanTicker == 1) {
	  $scan{$dataKey} = $dataValue;
	  $scanTicker = 0;
	} else {

	  if ($kscLookup{$biome[-1]}) {
	    next if $opt{noksc};
	    # KSC biomes *should* be SrfLanded-only, this ensures that we skip any
	    # anomalous data in persistent.sfs.  This complements the test below
	    # but saves some work given the KSC/Kerbin potential with the -k flag
	    next if $where[-1] ne 'Landed';
	    # Take KSC out of Kerbin
	    if (!$opt{ksckerbin}) {
	      $spob[-1] = $ksc;
	      $dataKey = $ksc.$where[-1];
	      ${$dataValue}[1] = $ksc;
	    }
	  }
	  # Stock science gets bigger keys and gets the biome added in the value
	  $dataKey = $test[-1].$dataKey.$biome[-1];

	  # Skip over annoying "fake" science expts caused by ScienceAlert, etc.
	  # For more info see
	  # http://forum.kerbalspaceprogram.com/threads/76793-0-90-ScienceAlert-1-8-4-Experiment-availability-feedback-%28December-23%29?p=1671187&viewfull=1#post1671187
	  # Still present (especially @ KSC (see above) but not much elsewhere).
	  # Has the annoying side effect of removing any science that *belongs*
	  # that I haven't manually included (alternate KSC sites, for example).
	  # Would presumably mean issues with science from other mods.  Would be
	  # better if I could explicitly remove the clearly wrong things and then
	  # just include anything I haven't accounted for, without also limiting
	  # things we've intentionally left out (e.g. infraredTelescope).  One
	  # idea is to check each item in the persistent.sfs field against the
	  # testdefs (for science we don't know about) and biomes against places
	  # we might not know about (although that one seems more tedious than is
	  # worthwhile).  If there's something we don't know about, and it's not
	  # 0, maybe store it and report?  It would help if I inserted recovery
	  # and SCANsat into testdefs (well, or properly process SCANsat's scidefs
	  # file) FIXME TODO
	  next if !$dataMatrix{$dataKey};

	  # Should note what and why this is here, I think to insert right biome
	  # FIXME TODO
	  splice $dataValue->@*, 3, 0, $biome[-1];
	  $dataMatrix{$dataKey} = $dataValue;
	}
      }
    }
  }
}

sub calcPerc {
  my ($sciC, $capC) = @_;
  return sprintf '%.2f', 100 * $sciC / $capC;
}

# Determine column header widths
sub columnWidths {
  my ($sheet, $colRef) = @_;

  foreach (0 .. $#{$colRef}) {
    $sheet->set_column($_, $_, ${$colRef}[$_]);
  }

  return;
}
# Build report data
sub processData {
  my ($dataRef, $shortName, $longName) = @_;

  foreach my $key (sort {specialSort($a, $b, $dataRef)} keys $dataRef->%*) {
    dataSplice($dataRef->{$key})                               if !$opt{moredata};
    writeToExcel($shortName, $dataRef->{$key}, $key, $dataRef) if (!$opt{excludeexcel} && (!$opt{unfinishedonly} || $dataRef->{$key}[-1] ne '100.00'));
    writeToCSV($dataRef->{$key})                               if ($opt{csv} && (!$opt{unfinishedonly} || $dataRef->{$key}[-1] ne '100.00'));

    if ($opt{tests}) {
      # Neater spacing in test averages output
      buildScienceData($key, $longName, \%testData, $dataRef);
      if ($opt{report}) {
	buildReportData($key, $shortName, $dataRef->{$key}[0], $dataRef);
      }
    } elsif ($opt{average}) {
      buildScienceData($key, $shortName, \%spobData, $dataRef);
      if ($opt{report}) {
	buildReportData($key, $shortName, $dataRef->{$key}[1], $dataRef);
      }
    }
  }
}


## Custom sort order, adapted from:
## http://stackoverflow.com/a/8171591/2521092
# Things with no biomes (recovery, SCANsat) but planet-wide: Kerbin, KSC, and
# its moons come first, then Kerbol, then proper sorting of conditions
# matches worksheets (What does that mean?)
sub specialSort {
  my ($a, $b, $specRef) = @_;

  # Grab all the pieces we need from the inputs:
  ## v/w: situation (meaningless for recovery and actually the test for SCANsat)
  # Store lookup for each, as we don't need to repeat the regex each time we see
  # a given item.
  $memoized_situation{$a} //= $a =~ /($COND_RE)$/ ? $1 : undef;
  $memoized_situation{$b} //= $b =~ /($COND_RE)$/ ? $1 : undef;
  my ($v, $w) = ($memoized_situation{$a}, $memoized_situation{$b});

  # Percent done, test, situation/test
  if ($opt{percentdone}) {
    return ${$specRef}{$b}[9] <=> ${$specRef}{$a}[9] || $a cmp $b || $cond_order_map{$v} <=> $cond_order_map{$w};
  }
  # Science left, test, situation/test
  if ($opt{scienceleft}) {
    return ${$specRef}{$b}[8] <=> ${$specRef}{$a}[8] || $a cmp $b || $cond_order_map{$v} <=> $cond_order_map{$w};
  }
  ## x/y: spob
  # As above
  $memoized_spob{$a} //= $a =~ /^($SPOB_RE)/ ? $1 : undef;
  $memoized_spob{$b} //= $b =~ /^($SPOB_RE)/ ? $1 : undef;
  my ($x, $y) = ($memoized_spob{$a}, $memoized_spob{$b});
  # Spob, situation/test
  return $spec_order_map{$x} <=> $spec_order_map{$y} || $cond_order_map{$v} <=> $cond_order_map{$w};
}

# Sort alphabetically by test, then specifically by situation, then
# alphabetically by biome.  Should probably include spob order FIXME TODO
sub sitSort {
  # Grab all the pieces we need from the inputs:
  ## v/w: test (and spob)
  ## x/y: situation (in order)
  ## t/u: biome
  # As above, values stored to speed things up
  $memoized_situation{$a} //= $a =~ /^(.+)($SIT_RE)(.+)$/ ? [$1, $2, $3] : undef;
  $memoized_situation{$b} //= $b =~ /^(.+)($SIT_RE)(.+)$/ ? [$1, $2, $3] : undef;
  my ($v, $x, $t, $w, $y, $u) = ($memoized_situation{$a}->@*, $memoized_situation{$b}->@*);

  # Percent done, test, situation, biome
  if ($opt{percentdone}) {
    return $dataMatrix{$b}[10] <=> $dataMatrix{$a}[10] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
  }
  # Science left, test, situation, biome
  if ($opt{scienceleft}) {
    return $dataMatrix{$b}[9] <=> $dataMatrix{$a}[9] || $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
  }
  # Biome, situation, test
  if ($opt{biome}) {
    return $t cmp $u || $sit_order_map{$x} <=> $sit_order_map{$y} || $v cmp $w;
  }
  # Test, situation, biome
  return $v cmp $w || $sit_order_map{$x} <=> $sit_order_map{$y} || $t cmp $u;
}


# Properly splice data
sub dataSplice {
  my $rowRef = shift;
  splice @{$rowRef}, 3, 3;
  return;
}

sub writeToCSV {
  my $rowRef = shift;

  print $csvOut join q{,}, @{$rowRef};
  print $csvOut "\n";
  return;
}

sub writeToExcel {
  my ($sheetName, $rowRef, $matrixKey, $hashRef) = @_;

  $workVars{$sheetName}[0]->write_row($workVars{$sheetName}[1], 0, $rowRef);
  $workVars{$sheetName}[0]->write($workVars{$sheetName}[1], $dataIdx, ${$hashRef}{$matrixKey}[$dataIdx], $bgRed) if ${$hashRef}{$matrixKey}[$dataIdx + 1] < $colorThreshold;

  if ($opt{moredata}) {
    $workVars{$sheetName}[0]->write($workVars{$sheetName}[1], 4, ${$hashRef}{$matrixKey}[4], $bgGreen) if ((${$hashRef}{$matrixKey}[4] < 0.001) && (${$hashRef}{$matrixKey}[4] > 0));
  }

  $workVars{$sheetName}[1]++;
  return;
}

# Build data hashes for averages
sub buildScienceData {
  my ($key, $ind, $dataRef, $hashRef) = @_;

  # Sci, count, cap
  ${$dataRef}{$ind}[0] += ${$hashRef}{$key}[$dataIdx];
  ${$dataRef}{$ind}[2] += ${$hashRef}{$key}[$dataIdx - 1];
  ${$dataRef}{$ind}[1]++;

  return;
}

# Build report
sub buildReportData {
  my ($key, $spo, $tes, $hashRef) = @_;
  $report{$spo}{$tes}     += ${$hashRef}{$key}[$dataIdx];
  $report{$spo}{$total}   += ${$hashRef}{$key}[$dataIdx];
  $report{$total}{$tes}   += ${$hashRef}{$key}[$dataIdx];
  $report{$total}{$total} += ${$hashRef}{$key}[$dataIdx];
  return;
}

# Alphabeticalish averages
sub averageAlphabetical {
  my $hashRef  = shift;
  my $arrayRef = shift;

  if ($opt{tests}) {
    push @{$arrayRef}, $recovery;                      # Neater spacing in test averages output
    push @{$arrayRef}, $scansatMap if $opt{scansat};
    @{$arrayRef} = sort @{$arrayRef};
  }

  foreach my $index (0 .. $#{$arrayRef}) {
    printAverageTable(${$arrayRef}[$index], $hashRef);
  }

  if (!$opt{tests}) {
    printAverageTable($recov,   $hashRef);
    printAverageTable($scansat, $hashRef) if $opt{scansat};
  }

  return;
}

# Averages sorted by total remaining science
sub averageRemaining {
  my $hashRef = shift;

  foreach my $key (sort {${$hashRef}{$b}[0] <=> ${$hashRef}{$a}[0] || $a cmp $b} keys %{$hashRef}) {
    printAverageTable($key, $hashRef);
  }

  return;
}

# Averages sorted by percent accomplished
sub averagePercent {
  my $hashRef = shift;

  foreach my $key (sort {((${$hashRef}{$b}[2] - ${$hashRef}{$b}[0]) / ${$hashRef}{$b}[2]) <=> ((${$hashRef}{$a}[2] - ${$hashRef}{$a}[0]) / ${$hashRef}{$a}[2]) || $a cmp $b} keys %{$hashRef}) {
    printAverageTable($key, $hashRef);
  }

  return;
}

# Handle printing of the averages table
sub printAverageTable {

  my @placeHolder = @_;
  my $ind         = $placeHolder[0];
  my %hash        = %{$placeHolder[1]};

  my $indShort = substr $ind, 0, 14;    # Neater spacing in test averages output
  my $avg      = $hash{$ind}[0] / ($hash{$ind}[1]);
  my $remains  = $hash{$ind}[2] - $hash{$ind}[0];
  my $per      = 100 * $remains / $hash{$ind}[2];

  printf "%s\t%.0f\t%.0f\t%.0f\n", $indShort, $avg, $hash{$ind}[0], $per;
  if ($opt{outputavgtable}) {
    printf $avgOut "%s\t%.0f\t%.0f\t%.0f\n", $ind, $avg, $hash{$ind}[0], $per;
  }

  return;
}

# Handle printing of the report table
sub printReportTable {
  my @placeHolder = @_;
  print $rptOut 'spob,';

  foreach my $place (sort @placeHolder) {
    print $rptOut "$place,";
  }
  print $rptOut "$total\n";

  foreach my $key (sort keys %report) {
    print $rptOut "$key";
    foreach my $subj (sort keys %{$report{'Kerbin'}}) {
      print $rptOut q{,};
      if ($report{$key}{$subj}) {
	printf $rptOut '%.0f', $report{$key}{$subj};
      }
    }
    print $rptOut "\n";
  }

  return;
}

#### Usage statement ####
# Escapes not necessary but ensure pretty colors
# Final line must be unindented?
sub usage {
  print <<"USAGE";
Usage: $PROGRAM_NAME [-atbspijdklmzcneor -h -f path/to/dotfile ]
       $PROGRAM_NAME [-g <game_location> -u <savefile_name>]

       $PROGRAM_NAME [-ATBSPIJDKLMZCNEOR -G -U] -> Turn off a given option

      -a Display average science left for each planet
      -t Display average science left for each experiment type.  Supersedes -a.

      -b Sort by biome, only output data file(s).
      -s Sort by science left, including output file(s) and averages from -a
         and -t flags.  Supersedes -b.
      -p Sort by percent science accomplished, including output file(s) and
         averages from -a and -t flags.  Supersedes -b and -s.

      -i Include data from SCANsat
      -j Ignore and don't consider asteroids or comets.
      -d Include science from the Breaking Ground expansion.
      -k List data from KSC biomes as being from Kerbin (in the same Excel worksheet)
      -l Ignore science for KSC biomes entirely
      -m Add some largely boring data to the output (i.e., dsc, sbv, scv)
      -z Don't include finished experiments, only those with science remaining
      -c Output data to csv file as well
      -n Turn off formatted printing in Excel (i.e., colors and bolding)
      -e Don't output the Excel file
      -o Save the chosen average table to a file.  Requires -a or -t.
      -r Save a report csv of per-planet condition or test data.  Require -a or -t.

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
