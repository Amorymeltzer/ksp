#!/usr/bin/env perl
# boundaries.pl by Amory Meltzer
# https://github.com/Amorymeltzer/ksp
# Print condition boundaries (InSpaceHigh, etc.)
# Data from wiki and http://forum.kerbalspaceprogram.com/threads/53567


use strict;
use warnings;
use diagnostics;

my %boundData;
# Hacky, but avoids an error below
my @universe = qw (Kerbin@FlyingHigh);

# Construct deltaV hash
while (<DATA>) {
  chomp;
  my @borders = split;
  my $key = $borders[0].q{@}.$borders[1];

  $boundData{$key} = $borders[2];

  # Build sorted list for reference when sorting later
  @universe = (@universe, $key) if $key ne $universe[-1];
}

print "Spob\tCondition\tAltitude (km))\n";

foreach my $key (sort { specialSort($a, $b) } keys %boundData) {
  my @tmp = split /@/, $key;
  print "$tmp[0]\t$tmp[1]\t$boundData{$key}\n";
}


# Sort via order in _DATA_
sub specialSort
  {
    my ($a,$b) = @_;
    my @input = ($a, $b);	# Keep 'em separate, avoid expr version of map

    my @specOrder = @universe;
    my %spec_order_map = map { $specOrder[$_] => $_ } 0 .. $#specOrder;
    my $sord = join q{|}, @specOrder;

    my ($x,$y) = map {/^($sord)/} @input;

    $spec_order_map{$x} <=> $spec_order_map{$y};
  }


## The lines below do not represent Perl code, and are not examined by the
## compiler.  Rather, they are the default boundary heights between different
## conditions for each space object
__END__
Kerbin FlyingHigh 18
  Kerbin InSpaceLow 69
  Kerbin InSpaceHigh 250
  Mun InSpaceHigh 60
  Minmus InSpaceHigh 30
  Kerbol InSpaceHigh 1000000
  Moho InSpaceHigh 80
  Eve FlyingHigh 22
  Eve InSpaceLow 60?
  Eve InSpaceHigh 400
  Gilly InSpaceHigh 6
  Duna FlyingHigh 12
  Duna InSpaceLow ???
  Duna InSpaceHigh 140
  Ike InSpaceHigh 50
  Dres InSpaceHigh 25
  Jool FlyingHigh 120?
  Jool InSpaceLow 200?
  Jool InSpaceHigh 4000
  Laythe FlyingHigh 10?
  Laythe InSpaceLow 15?
  Laythe InSpaceHigh 200
  Vall InSpaceHigh 90
  Tylo InSpaceHigh 250
  Bop InSpaceHigh 25
  Pol InSpaceHigh 22
  Eeloo InSpaceHigh 60
