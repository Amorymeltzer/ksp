#!/usr/bin/env perl
# boundaries.pl by Amory Meltzer
# Print condition boundaries (InSpaceHigh, etc.)

use strict;
use warnings;
use diagnostics;

my %boundData;			# Hold delta-V values from END data


# Construct deltaV hash
while (<DATA>) {
  chomp;
  my @borders = split;
  $boundData{$borders[0]."@".$borders[1]} = $borders[2];
}

print "Spob\tCondition\tAltitude\n";

foreach my $key (keys %boundData) {
  my @tmp = split /@/, $key;
  print "$tmp[0]\t$tmp[1]\t$boundData{$key}\n";
}

## The lines below do not represent Perl code, and are not examined by the
## compiler.  Rather, they are the default boundary heights between different
## conditions for each space object
__END__
Kerbol InSpaceHigh 1000000
  Kerbin FlyingHigh 18
  Kerbin InSpaceLow 70
  Kerbin InSpaceHigh 250
  Mun InSpaceHigh 60
  Minmus InSpaceHigh 30
  Moho InSpaceHigh 80
  Eve FlyingHigh 22
  Eve InSpaceLow ???
  Eve InSpaceHigh 400
  Gilly InSpaceHigh 6
  Duna FlyingHigh 12
  Duna InSpaceLow ???
  Duna InSpaceHigh 140
  Ike InSpaceHigh 50
  Dres InSpaceHigh 25
  Jool FlyingHigh 120
  Jool InSpaceLow ???
  Jool InSpaceHigh 4000
  Laythe FlyingHigh 10
  Laythe InSpaceLow ???
  Laythe InSpaceHigh 200
  Vall InSpaceHigh 90
  Tylo InSpaceHigh 250
  Bop InSpaceHigh 25
  Pol InSpaceHigh 22
  Eeloo InSpaceHigh 60
