#!/usr/bin/env perl
# deltaVScience.pl by Amory Meltzer
# https://github.com/Amorymeltzer/ksp
# Attempt to (roughly) estimate a ratio of science per deltaV for each spob
# Takes in an average (-a or -as) table from parseScience.pl
# Delta-V from http://mononyk.us/wherecanigo.php?dv=0&loc=orbit&figs=diff
## Use full output for situation-specific calculations?
## Merge recovery into planetary spob?  Eh, maybe doesn't make sense.

use strict;
use warnings;
use diagnostics;


my $avgData = 'average_table.txt';
my %deltaVData;			# Hold delta-V values from END data


# Construct deltaV hash
while (<DATA>) {
  chomp;
  my @dvs = split;
  $deltaVData{$dvs[0]} = $dvs[1];
}

print "Spob\tdV avg\tdV tot\n";
# Read in averages
open my $avg, '<', "$avgData" or die $!;
while (<$avg>) {
  next until $. > 3;		# Skip header lines
  chomp;
  # No delta-V for recovery, Kerbin, KSC, SCANsat
  next if /^R|^Kerbi|^KSC|^SCANsat/;
  my @avgs = split /\t/;

  # Is this the best sclaing factor?
  my $effA = 1000*$avgs[1]/$deltaVData{$avgs[0]};
  my $effT = 1000*$avgs[2]/$deltaVData{$avgs[0]};

  printf "%s\t%0.f\t%0.f\n", $avgs[0], $effA, $effT;
}
close $avg or die $!;







## The lines below do not represent Perl code, and are not examined by the
## compiler.  Rather, they are estimated deltaV requirements to the lowest
## accesible spot on each spob from the surface of Kerbin.  In the future this
## may become situation specific, as that would be overwhelmingly more
## exact. Right now Jool is to the surface, Eve is landing.
__END__
Mun 6260
  Minmus 5790
  Kerbol 38230
  Moho 9830
  Eve 18890
  Gilly 8785
  Duna 7360
  Ike 6895
  Dres 7205
  Jool 9095
  Laythe 14275
  Vall 13035
  Tylo 15295
  Bop 12751
  Pol 12675
  Eeloo 9590
